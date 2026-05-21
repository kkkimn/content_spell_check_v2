"""
core_video.py — 정확도 + 속도 + 안정성 업그레이드 버전

[정확도]
  1. Whisper `prompt` 도메인 힌트 주입
  2. 세그먼트 문장 단위 병합 (이중피동·어미 혼동 검출률↑)
  3. STT 오인식 가이드 + Few-shot + JSON Schema structured output
  4. OCR pHash 중복 제거

[속도]
  1. OCR 배치 병렬 호출 (ThreadPoolExecutor) — 2~5배 단축
  2. 프레임 추출 `cap.grab()` + 조건부 `retrieve()` 패턴
  3. ffmpeg 직접 호출 오디오 추출 (미설치 시 moviepy 폴백)
  4. JPEG 인코딩 품질/해상도 튜닝
  5. 음성·화면 파이프라인 병렬 실행 헬퍼

[안정성] ★ 이번 업그레이드
  1. 25MB 초과 오디오 자동 청크 분할 — 긴 영상 안전 처리
     + 청크별 병렬 STT + 원본 기준 타임스탬프 재조정
  2. Rate limit (429) 전용 처리 — Retry-After 헤더 존중, 더 긴 백오프
  3. STT 결과 캐싱 — 파일 해시 기반 (재실행 비용 0)
  4. 중간 실패 격리 — 한 배치 실패해도 나머지 유지 (이미 적용됨)
  5. 임시 파일 안전 관리 — context manager 제공
"""

import os
import re
import io
import json
import time
import base64
import hashlib
import shutil
import tempfile
import subprocess
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import List, Optional, Callable, Dict, Any, Tuple

import cv2
import numpy as np
from moviepy.editor import VideoFileClip
from openai import OpenAI


# ─────────────────────────────────────────────
# 유틸리티
# ─────────────────────────────────────────────

_FFMPEG_BIN = shutil.which("ffmpeg")


def extract_audio(video_path, audio_path, sample_rate: int = 16000, mono: bool = True):
    """
    MP4 영상에서 오디오(.mp3)를 추출합니다.

    속도 우선:
    - ffmpeg가 설치돼 있으면 subprocess로 직접 호출 (moviepy 대비 3~5배 빠름)
    - 16kHz / mono 다운샘플로 업로드 크기 축소 + Whisper 정확도는 거의 영향 없음
    - ffmpeg 미설치 시 moviepy로 폴백

    Parameters
    ----------
    sample_rate : int
        오디오 샘플링 레이트 (Hz). Whisper는 16kHz로 내부 리샘플하므로 16000 권장.
    mono : bool
        모노 다운믹스 여부. 강의/대화 영상은 모노로 충분.
    """
    # 1) ffmpeg 경로가 있으면 직접 호출
    if _FFMPEG_BIN:
        try:
            cmd = [
                _FFMPEG_BIN,
                "-y",                     # 덮어쓰기
                "-loglevel", "error",
                "-i", video_path,
                "-vn",                    # 비디오 제거
                "-ar", str(sample_rate),
                "-ac", "1" if mono else "2",
                "-b:a", "64k",            # 64kbps 충분
                audio_path,
            ]
            subprocess.run(cmd, check=True, capture_output=True)
            return True
        except subprocess.CalledProcessError as e:
            stderr = (e.stderr or b"").decode("utf-8", errors="ignore")
            print(f"[ffmpeg 실패, moviepy로 폴백] {stderr.strip()[:200]}")
        except Exception as e:
            print(f"[ffmpeg 예외, moviepy로 폴백] {e}")

    # 2) 폴백: moviepy
    try:
        video = VideoFileClip(video_path)
        video.audio.write_audiofile(
            audio_path,
            fps=sample_rate,
            nbytes=2,
            codec="libmp3lame",
            bitrate="64k",
            ffmpeg_params=["-ac", "1" if mono else "2"],
            logger=None,
        )
        video.close()
        return True
    except Exception as e:
        print(f"오디오 추출 오류: {e}")
        return False


def format_timestamp(seconds):
    """초(float) 단위의 시간을 [HH:MM:SS] 또는 [MM:SS] 형태로 변환합니다."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    if hours > 0:
        return f"[{hours:02d}:{minutes:02d}:{secs:02d}]"
    return f"[{minutes:02d}:{secs:02d}]"


def call_with_retry(fn, retries=4, delay=3, max_delay: float = 60.0):
    """
    API 호출 실패 시 exponential backoff으로 재시도합니다.

    ★ 안정성 업그레이드:
    - RateLimitError(429) 감지 시 Retry-After 헤더를 존중 (있으면 그대로, 없으면 더 긴 대기)
    - APIConnectionError / Timeout 계열도 지수 백오프로 재시도
    - 최대 대기 시간(max_delay)으로 상한 둠

    retries : 총 시도 횟수 (기본 4)
    delay   : 초기 지연(초). 이후 2배씩 증가.
    """
    for attempt in range(retries):
        try:
            return fn()
        except Exception as e:
            err_name = type(e).__name__
            msg = str(e)
            is_rate_limit = "RateLimit" in err_name or "429" in msg

            # Retry-After 헤더 파싱 시도
            retry_after: Optional[float] = None
            try:
                resp = getattr(e, "response", None)
                if resp is not None:
                    hdr = getattr(resp, "headers", {}) or {}
                    ra = hdr.get("Retry-After") or hdr.get("retry-after")
                    if ra:
                        retry_after = float(ra)
            except Exception:
                pass

            if attempt < retries - 1:
                if is_rate_limit:
                    wait = retry_after if retry_after else min(max_delay, delay * (2 ** attempt) * 2)
                    print(f"⚠️ Rate limit ({err_name}) — {wait:.1f}초 대기 후 재시도 "
                          f"(시도 {attempt + 1}/{retries})")
                else:
                    wait = min(max_delay, delay * (2 ** attempt))
                    print(f"API 호출 오류 (시도 {attempt + 1}/{retries}) [{err_name}]: {msg[:200]}")
                time.sleep(wait)
            else:
                print(f"API 호출 최종 실패 [{err_name}]: {msg[:200]}")
    raise RuntimeError(f"API 호출이 {retries}회 모두 실패했습니다.")


# ─────────────────────────────────────────────
# 오디오 청크 분할 유틸 (25MB 초과 안전 처리)
# ─────────────────────────────────────────────

WHISPER_FILE_LIMIT_BYTES = 25 * 1024 * 1024   # 25 MB (Whisper API 제한)
# 실제로는 여유를 두고 20MB 기준으로 분할 판단
_CHUNK_SIZE_THRESHOLD = 20 * 1024 * 1024


def _get_audio_duration(audio_path: str) -> Optional[float]:
    """ffprobe로 오디오 길이(초)를 구합니다. 실패 시 None."""
    ffprobe = shutil.which("ffprobe")
    if not ffprobe:
        return None
    try:
        result = subprocess.run(
            [ffprobe, "-v", "error", "-show_entries", "format=duration",
             "-of", "default=noprint_wrappers=1:nokey=1", audio_path],
            capture_output=True, text=True, check=True, timeout=10,
        )
        return float(result.stdout.strip())
    except Exception:
        return None


def _split_audio_into_chunks(
    audio_path: str,
    chunk_duration: float,
    output_dir: str,
) -> List[Tuple[str, float]]:
    """
    ffmpeg로 오디오를 시간 기반으로 분할합니다.

    Returns
    -------
    List[(chunk_path, chunk_start_seconds)]
        각 청크 파일 경로와 원본 기준 시작 시각.
    """
    if not _FFMPEG_BIN:
        raise RuntimeError("오디오 청크 분할에는 ffmpeg가 필요합니다.")

    total_duration = _get_audio_duration(audio_path)
    if total_duration is None:
        # ffprobe 없어도 동작하게: 고정 개수로 나눌 수 없으므로 에러
        raise RuntimeError("ffprobe를 찾을 수 없어 오디오 길이를 확인할 수 없습니다.")

    chunks: List[Tuple[str, float]] = []
    idx = 0
    start = 0.0
    while start < total_duration:
        chunk_path = os.path.join(output_dir, f"chunk_{idx:03d}.mp3")
        # -ss는 입력 앞에 두면 빠른 seek, 뒤에 두면 정확한 seek.
        # 정확도를 위해 출력 옵션으로 사용.
        cmd = [
            _FFMPEG_BIN, "-y", "-loglevel", "error",
            "-i", audio_path,
            "-ss", f"{start:.3f}",
            "-t", f"{chunk_duration:.3f}",
            "-c", "copy",   # 재인코딩 없이 빠르게 복사
            chunk_path,
        ]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=60)
        except subprocess.CalledProcessError:
            # copy 실패 시 재인코딩으로 폴백
            cmd_reenc = [
                _FFMPEG_BIN, "-y", "-loglevel", "error",
                "-i", audio_path,
                "-ss", f"{start:.3f}",
                "-t", f"{chunk_duration:.3f}",
                "-ar", "16000", "-ac", "1", "-b:a", "64k",
                chunk_path,
            ]
            subprocess.run(cmd_reenc, check=True, capture_output=True, timeout=120)

        if os.path.exists(chunk_path) and os.path.getsize(chunk_path) > 0:
            chunks.append((chunk_path, start))
        idx += 1
        start += chunk_duration

    return chunks


def _plan_audio_chunks(audio_path: str, target_chunk_bytes: int = _CHUNK_SIZE_THRESHOLD) -> float:
    """
    파일 크기와 길이로부터 **청크당 적정 시간(초)** 을 계산합니다.
    """
    size = os.path.getsize(audio_path)
    duration = _get_audio_duration(audio_path) or 0.0
    if duration <= 0:
        # 안전 기본값: 10분
        return 600.0
    # 대략 "(목표_바이트 / 전체_바이트) * 전체_초"
    ratio = target_chunk_bytes / max(size, 1)
    chunk_sec = max(60.0, min(1500.0, duration * ratio * 0.95))   # 최소 1분, 최대 25분
    return chunk_sec


@contextmanager
def temp_workspace(prefix: str = "spellcheck_"):
    """
    임시 작업 디렉토리를 만들고, 블록 종료 시 자동 정리합니다.

    사용:
        with temp_workspace() as tmpdir:
            ...
    """
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    try:
        yield tmpdir
    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass


# ─────────────────────────────────────────────
# 음성(STT) 처리
# ─────────────────────────────────────────────

@dataclass
class MergedSegment:
    """문장 단위로 병합된 세그먼트."""
    id: int
    start: float
    end: float
    text: str
    # 병합에 포함된 원본 Whisper 세그먼트 인덱스
    source_ids: List[int] = field(default_factory=list)


# 한국어 문장 종결 패턴 — 문장 병합 경계 판별용
_SENTENCE_END_PATTERN = re.compile(
    r"[.!?…]$|"
    r"(다|요|죠|네|까|군|구나|라|까요|어요|아요|예요|이에요|이다|입니다|습니다|니다)[.!?]?$"
)


def _is_sentence_end(text: str) -> bool:
    """문장이 끝났는지 추정 (한국어 종결어미 또는 문장부호)."""
    t = text.strip()
    if not t:
        return False
    return bool(_SENTENCE_END_PATTERN.search(t))


def merge_segments_into_sentences(
    segments,
    max_gap: float = 1.5,
    max_duration: float = 15.0,
    min_chars: int = 12,
):
    """
    Whisper의 짧은 세그먼트를 문장 단위로 병합합니다.

    병합 중단 조건:
    - 이전 세그먼트 끝과 현재 세그먼트 시작 간격(gap) > max_gap 초
    - 누적 길이(duration) >= max_duration 초
    - 이전 세그먼트가 문장 종결 형태이고 누적 길이가 min_chars 이상

    Parameters
    ----------
    segments : list
        Whisper verbose_json의 segments (obj 또는 dict)
    max_gap : float
        이 이상 침묵이면 강제로 문장 분리
    max_duration : float
        이 이상 길어지면 강제로 문장 분리 (너무 긴 입력 방지)
    min_chars : int
        최소 이 글자 수가 넘어야 문장 종료로 판정

    Returns
    -------
    List[MergedSegment]
    """
    def _get(seg, key):
        return getattr(seg, key, None) if not isinstance(seg, dict) else seg.get(key)

    merged: List[MergedSegment] = []
    buf_start: Optional[float] = None
    buf_end: Optional[float] = None
    buf_text = ""
    buf_sources: List[int] = []

    def _flush():
        nonlocal buf_start, buf_end, buf_text, buf_sources
        if buf_text.strip():
            merged.append(MergedSegment(
                id=len(merged),
                start=buf_start if buf_start is not None else 0.0,
                end=buf_end if buf_end is not None else 0.0,
                text=buf_text.strip(),
                source_ids=buf_sources.copy(),
            ))
        buf_start, buf_end, buf_text, buf_sources = None, None, "", []

    for i, seg in enumerate(segments):
        text = (_get(seg, "text") or "").strip()
        if not text:
            continue
        # Whisper가 뱉는 "." 만, "음…" 같은 무의미한 세그먼트 필터
        if len(text) <= 1 and not text.isalnum():
            continue

        start = float(_get(seg, "start") or 0.0)
        end = float(_get(seg, "end") or start)

        # 첫 버퍼 초기화
        if buf_start is None:
            buf_start, buf_end = start, end
            buf_text = text
            buf_sources = [i]
            continue

        gap = start - buf_end
        if gap > max_gap:
            _flush()
            buf_start, buf_end = start, end
            buf_text = text
            buf_sources = [i]
            continue

        # 누적 길이 체크
        if (end - buf_start) > max_duration:
            _flush()
            buf_start, buf_end = start, end
            buf_text = text
            buf_sources = [i]
            continue

        # 이어붙이기
        buf_text = f"{buf_text} {text}".strip()
        buf_end = end
        buf_sources.append(i)

        # 문장 종결 판정
        if _is_sentence_end(buf_text) and len(buf_text) >= min_chars:
            _flush()

    _flush()
    return merged


def _transcribe_single_file(client: OpenAI, audio_path: str, prompt: str):
    """단일 오디오 파일 전사. 내부용."""
    def _call():
        with open(audio_path, "rb") as audio_file:
            return client.audio.transcriptions.create(
                model="whisper-1",
                file=audio_file,
                language="ko",
                response_format="verbose_json",
                timestamp_granularities=["segment", "word"],
                prompt=prompt,
                temperature=0.0,
            )
    return call_with_retry(_call)


def _shift_segments(segments, offset: float):
    """세그먼트의 start/end에 offset(초)를 더해 원본 타임라인으로 복원."""
    shifted = []
    for seg in segments:
        is_dict = isinstance(seg, dict)
        start = (seg.get("start") if is_dict else getattr(seg, "start", 0)) or 0
        end   = (seg.get("end")   if is_dict else getattr(seg, "end",   0)) or 0
        text  = (seg.get("text")  if is_dict else getattr(seg, "text", "")) or ""
        shifted.append(_ShiftedSeg(
            start=float(start) + offset,
            end=float(end) + offset,
            text=text,
        ))
    return shifted


@dataclass
class _ShiftedSeg:
    """청크 병합용 경량 세그먼트 (merge_segments_into_sentences와 호환)."""
    start: float
    end: float
    text: str


def transcribe_audio(
    audio_path,
    api_key,
    model: str = "whisper-1",
    domain_hint: str = "",
    max_chunk_bytes: int = _CHUNK_SIZE_THRESHOLD,
    max_workers: int = 3,
):
    """
    오디오 → 텍스트 변환.

    ★ 안정성 업그레이드:
    - 파일이 Whisper 25MB 제한을 초과하면 자동으로 ffmpeg로 청크 분할 후 병렬 전송
    - 각 청크 결과의 타임스탬프를 원본 기준으로 재조정해 이어 붙임
    - 한 청크가 실패해도 나머지 청크 결과는 보존 (부분 복구)

    Parameters
    ----------
    max_chunk_bytes : int
        청크 분할 임계값. 기본 20MB (25MB 한도에 여유 둠).
    max_workers : int
        청크 병렬 전사 워커 수. 권장 2~4.

    Returns
    -------
    segments : Whisper verbose_json segments (원본 타임라인으로 복원됨)
    """
    client = OpenAI(api_key=api_key)

    if model != "whisper-1":
        print(f"[안내] {model}은 세그먼트 타임스탬프를 지원하지 않아 "
              f"whisper-1로 자동 폴백합니다. (정확도 보완은 프롬프트 힌트로 유지)")

    prompt = (
        "다음은 한국어 강의 또는 발표 영상의 음성입니다. "
        "정확한 한국어 맞춤법과 고유명사 표기를 사용해 전사하세요."
    )
    if domain_hint.strip():
        prompt += f" 주요 용어: {domain_hint.strip()}"

    # ── 파일 크기 확인 ─────────────────────────────
    try:
        file_size = os.path.getsize(audio_path)
    except OSError as e:
        raise RuntimeError(f"오디오 파일을 읽을 수 없습니다: {e}")

    # 25MB 이하면 단일 호출
    if file_size <= max_chunk_bytes:
        transcription = _transcribe_single_file(client, audio_path, prompt)
        return transcription.segments

    # ── 25MB 초과: 청크 분할 ───────────────────────
    size_mb = file_size / (1024 * 1024)
    print(f"📦 오디오 크기 {size_mb:.1f}MB → Whisper 25MB 한도 초과, 청크 분할 시작")

    if not _FFMPEG_BIN:
        raise RuntimeError(
            f"오디오가 {size_mb:.1f}MB로 Whisper 25MB 한도를 초과합니다. "
            f"청크 분할을 위해 ffmpeg 설치가 필요합니다."
        )

    chunk_seconds = _plan_audio_chunks(audio_path, max_chunk_bytes)
    print(f"   → 청크 길이: {chunk_seconds:.0f}초 (약 {chunk_seconds/60:.1f}분)")

    all_segments: List[Any] = []
    with temp_workspace(prefix="stt_chunks_") as tmpdir:
        try:
            chunks = _split_audio_into_chunks(audio_path, chunk_seconds, tmpdir)
        except Exception as e:
            raise RuntimeError(f"오디오 청크 분할 실패: {e}")

        if not chunks:
            raise RuntimeError("오디오 청크 분할 결과가 비어 있습니다.")

        print(f"   → 총 {len(chunks)}개 청크 생성, 병렬 전사 시작 (workers={max_workers})")

        def _process_chunk(chunk_path: str, offset: float, idx: int):
            try:
                tr = _transcribe_single_file(client, chunk_path, prompt)
                shifted = _shift_segments(tr.segments, offset)
                return idx, shifted, None
            except Exception as e:
                return idx, [], str(e)

        indexed_results: Dict[int, List[Any]] = {}
        errors: List[str] = []

        if max_workers <= 1 or len(chunks) == 1:
            for i, (cp, off) in enumerate(chunks):
                idx, segs, err = _process_chunk(cp, off, i)
                indexed_results[idx] = segs
                if err:
                    errors.append(f"청크 {idx+1}: {err}")
        else:
            with ThreadPoolExecutor(max_workers=max_workers) as ex:
                futures = {
                    ex.submit(_process_chunk, cp, off, i): i
                    for i, (cp, off) in enumerate(chunks)
                }
                for fut in as_completed(futures):
                    idx, segs, err = fut.result()
                    indexed_results[idx] = segs
                    if err:
                        errors.append(f"청크 {idx+1}: {err}")

        # 시간 순서대로 이어붙임
        for i in sorted(indexed_results.keys()):
            all_segments.extend(indexed_results[i])

    if errors:
        print(f"⚠️ 일부 청크 실패 ({len(errors)}개): {errors[0][:200]}")
        if not all_segments:
            raise RuntimeError(f"모든 청크 전사 실패: {errors[0]}")

    print(f"   ✅ 청크 전사 완료: {len(all_segments)}개 세그먼트")
    return all_segments


# ─────────────────────────────────────────────
# STT 결과 캐싱 (파일 해시 기반)
# ─────────────────────────────────────────────

def _file_content_hash(path: str, chunk_size: int = 1024 * 1024) -> str:
    """파일 SHA-256 해시. 큰 파일도 스트리밍으로 처리."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def compute_stt_cache_key(video_path: str, domain_hint: str = "") -> str:
    """
    STT 결과 캐싱용 키를 만듭니다.
    (파일 내용 해시 + 도메인 힌트)
    """
    file_hash = _file_content_hash(video_path)
    hint_hash = hashlib.md5(domain_hint.encode("utf-8")).hexdigest()[:8]
    return f"{file_hash[:16]}_{hint_hash}"


def save_stt_cache(cache_dir: str, cache_key: str, segments) -> bool:
    """전사 결과를 JSON으로 저장 (세그먼트는 start/end/text만 보존)."""
    try:
        os.makedirs(cache_dir, exist_ok=True)
        path = os.path.join(cache_dir, f"stt_{cache_key}.json")
        simplified = []
        for s in segments:
            is_dict = isinstance(s, dict)
            simplified.append({
                "start": float((s.get("start") if is_dict else getattr(s, "start", 0)) or 0),
                "end":   float((s.get("end")   if is_dict else getattr(s, "end",   0)) or 0),
                "text":  (s.get("text") if is_dict else getattr(s, "text", "")) or "",
            })
        with open(path, "w", encoding="utf-8") as f:
            json.dump(simplified, f, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"STT 캐시 저장 실패: {e}")
        return False


def load_stt_cache(cache_dir: str, cache_key: str):
    """캐시된 전사 결과를 로드. 없으면 None."""
    path = os.path.join(cache_dir, f"stt_{cache_key}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return [_ShiftedSeg(start=x["start"], end=x["end"], text=x["text"]) for x in raw]
    except Exception as e:
        print(f"STT 캐시 로드 실패: {e}")
        return None


# ─────────────────────────────────────────────
# 맞춤법 검사 — 공통 규칙 & 프롬프트
# ─────────────────────────────────────────────

_KO_SPELL_RULES = """
주요 검사 항목 (특히 집중):
1. 사이시옷: 숫자·뒷말 등 합성어 표기 (예: 나뭇잎, 숫자)
2. 된소리·거센소리 혼동: '깨끗이' vs '깨끗히', '않다' vs '안다'
3. 어미 혼동: '-이에요'/'-이어요', '-데'/'-대', '-든지'/'-던지'
4. 조사·의존명사 띄어쓰기: '것', '수', '때', '만큼', '뿐' 등
5. 피동·사동 표현 오용: '되어지다', '만들어지다' 등 이중피동
6. 외래어 표기: '컨텐츠→콘텐츠', '메세지→메시지' 등
7. 공식 명칭·고유어 오기: 틀린 한자어·혼용 표기
8. 불필요한 중복 표현: '미리 예방', '과반수 이상' 등
9. 단위·숫자 표기: '1 개월' 등 띄어쓰기
10. 어순 및 비문(문장 구조 오류)
"""

# STT 특성을 반영한 가이드 — false positive 감소
_STT_CAVEAT = """
【STT 변환 특성 주의사항】
이 텍스트는 음성인식(STT)으로 변환되어, 아래 경우에는 원문을 유지하세요:
- 발화자의 말버릇·구어체(예: "어", "음", "그니까")는 맞춤법 오류가 아님
- 외래어가 한글로 음차된 경우 문맥상 맞으면 유지
- 동음이의어 선택이 애매하면 원문 유지
- STT 경계가 어색해도 의미가 통하면 유지

반대로 **확실한 오류**는 반드시 교정:
- 명백한 표준어 규정 위반 (예: "됬다" → "됐다")
- 외래어 표기법 위반 (예: "컨텐츠" → "콘텐츠")
- 이중피동 (예: "되어지다" → "되다")
"""

_SPELL_FEWSHOT = """
【교정 예시】
원문: "영상 컨텐츠가 만들어지면 됩니다."
교정: "영상 <red>콘텐츠</red>가 <red>만들어지면</red> 됩니다."
      → 콘텐츠(외래어 표기), 만들어지면(이중피동 주의 — 문맥에 따라 '만들면')

원문: "해당 내용은 다음주에 발표할께요."
교정: "해당 내용은 <red>다음 주</red>에 <red>발표할게요</red>."
      → 다음 주(의존명사 띄어쓰기), 발표할게요(어미 '-ㄹ게')

원문: "과반수 이상이 찬성했어요."
교정: "<red>과반수가</red> 찬성했어요."
      → '과반수' 자체가 절반 초과 의미 → '이상' 중복 표현
"""

_SPELL_SYSTEM_PROMPT = f"""당신은 대한민국 방송 표준어 규정과 한글 맞춤법을 완벽히 숙지한 수석 교열 전문가입니다.
아래 텍스트는 동영상 음성을 STT로 변환한 결과입니다. (각 줄에 ID, 시간, 텍스트 포함)
일부 줄은 [CONTEXT] 표시가 되어 있으며, 이는 앞뒤 문맥 참고용으로만 사용하고 **교정 대상이 아닙니다**.

{_KO_SPELL_RULES}
{_STT_CAVEAT}
{_SPELL_FEWSHOT}

【작업 지시】
- [CONTEXT] 줄은 참고만 하고 결과에 포함하지 마세요.
- 전체 문맥을 파악한 뒤 교정하세요.
- 오류가 확실할 때만 교정. 애매하면 원문 유지.
- 수정된 단어·어절만 <red>단어</red>로 감싸세요.
- 원문 문장 전체를 반환하되 수정 부분만 태그 처리.
- 오류 없는 세그먼트는 결과 배열에서 제외.
"""

# Structured Output용 JSON Schema
_SPELL_SCHEMA = {
    "name": "spell_corrections",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "corrections": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "id":        {"type": "integer"},
                        "original":  {"type": "string"},
                        "corrected": {"type": "string"},
                        "reason":    {"type": "string"},
                    },
                    "required": ["id", "original", "corrected", "reason"],
                },
            }
        },
        "required": ["corrections"],
    },
}


# ─────────────────────────────────────────────
# 음성 맞춤법 검사 (세그먼트 병합 + 슬라이딩 윈도우)
# ─────────────────────────────────────────────

def spell_check_segments(
    segments,
    api_key,
    context_window: int = 2,
    model: str = "gpt-5.4",
    use_sentence_merge: bool = True,
    batch_size: int = 40,
    max_workers: int = 3,
):
    """
    세그먼트를 문장 단위로 병합한 뒤 맞춤법 검사합니다.

    Parameters
    ----------
    context_window : int
        각 배치 앞/뒤로 덧붙일 참고 세그먼트 수 (교정 대상 아님)
    use_sentence_merge : bool
        True면 Whisper 세그먼트를 문장 단위로 병합 (정확도↑)
    batch_size : int
        한 번에 모델에 보낼 교정 대상 세그먼트 수
    max_workers : int
        ★ 속도 업그레이드: 배치 병렬 호출 개수. OpenAI rate limit 고려 권장 2~4.
    """
    client = OpenAI(api_key=api_key)

    # 1) 문장 단위 병합
    if use_sentence_merge:
        merged = merge_segments_into_sentences(segments)
    else:
        merged = [
            MergedSegment(
                id=i,
                start=float(getattr(s, "start", s.get("start", 0)) if hasattr(s, "start") or isinstance(s, dict) else 0),
                end=float(getattr(s, "end", s.get("end", 0)) if hasattr(s, "end") or isinstance(s, dict) else 0),
                text=(getattr(s, "text", s.get("text", "")) if hasattr(s, "text") or isinstance(s, dict) else "").strip(),
                source_ids=[i],
            )
            for i, s in enumerate(segments)
            if (getattr(s, "text", s.get("text", "")) if hasattr(s, "text") or isinstance(s, dict) else "").strip()
        ]

    if not merged:
        return []

    # 2) 배치 생성 (컨텍스트 포함 텍스트 미리 빌드)
    batches: List[Tuple[int, int, str]] = []  # (batch_start, batch_end, batch_text)
    for batch_start in range(0, len(merged), batch_size):
        batch_end = min(batch_start + batch_size, len(merged))
        ctx_start = max(0, batch_start - context_window)
        ctx_end = min(len(merged), batch_end + context_window)

        lines = []
        for i in range(ctx_start, ctx_end):
            seg = merged[i]
            tag = "[CONTEXT] " if (i < batch_start or i >= batch_end) else ""
            lines.append(f"{tag}{format_timestamp(seg.start)} (ID:{seg.id}): {seg.text}")
        batches.append((batch_start, batch_end, "\n".join(lines)))

    def _process_batch(batch_idx: int, bs: int, be: int, text: str) -> List[dict]:
        """단일 배치 처리 — 병렬 워커용."""
        def _call():
            return client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": _SPELL_SYSTEM_PROMPT},
                    {"role": "user", "content": text},
                ],
                temperature=0.0,
                response_format={"type": "json_schema", "json_schema": _SPELL_SCHEMA},
            )

        batch_results = []
        try:
            response = call_with_retry(_call)
            parsed = json.loads(response.choices[0].message.content)
            corrections = parsed.get("corrections", [])

            for item in corrections:
                idx = item.get("id")
                if idx is None:
                    continue
                seg = next((m for m in merged if m.id == idx), None)
                if seg is None:
                    continue
                orig = (item.get("original") or "").strip()
                corr = (item.get("corrected") or "").strip()
                if not (orig and corr and orig != corr):
                    continue
                batch_results.append({
                    "구분": "음성 대본",
                    "시간": format_timestamp(seg.start),
                    "수정 전": _red_to_html(orig),
                    "수정 후": _red_to_html(corr),
                    "교정 사유": item.get("reason") or "",
                })
        except Exception as e:
            print(f"음성 맞춤법 검사 오류 (배치 {batch_idx + 1}): {e}")
            try:
                batch_results.extend(_fallback_text_mode_spell_check(
                    client, model, merged, bs, be, context_window
                ))
            except Exception as e2:
                print(f"  폴백도 실패: {e2}")
        return batch_results

    # 3) 병렬 실행
    results: List[dict] = []
    if max_workers <= 1 or len(batches) == 1:
        for i, (bs, be, text) in enumerate(batches):
            results.extend(_process_batch(i, bs, be, text))
    else:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {
                ex.submit(_process_batch, i, bs, be, text): i
                for i, (bs, be, text) in enumerate(batches)
            }
            # 시간순 정렬을 위해 인덱스별로 받은 뒤 합치기
            indexed: Dict[int, List[dict]] = {}
            for fut in as_completed(futures):
                batch_idx = futures[fut]
                try:
                    indexed[batch_idx] = fut.result()
                except Exception as e:
                    print(f"병렬 배치 {batch_idx + 1} 실패: {e}")
                    indexed[batch_idx] = []
            for i in sorted(indexed.keys()):
                results.extend(indexed[i])

    return results


def _fallback_text_mode_spell_check(client, model, merged, batch_start, batch_end, context_window):
    """json_schema를 지원하지 않는 모델용 폴백 (기존 방식)."""
    ctx_start = max(0, batch_start - context_window)
    ctx_end = min(len(merged), batch_end + context_window)

    lines = []
    for i in range(ctx_start, ctx_end):
        seg = merged[i]
        tag = "[CONTEXT] " if (i < batch_start or i >= batch_end) else ""
        lines.append(f"{tag}{format_timestamp(seg.start)} (ID:{seg.id}): {seg.text}")

    instruction = (
        _SPELL_SYSTEM_PROMPT
        + "\n\n응답은 반드시 아래 순수 JSON 배열 형식으로만(백틱·설명 없이):\n"
        '[{"id":0,"original":"…","corrected":"…","reason":"…"}]\n오류 없으면 []'
    )

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": instruction},
            {"role": "user", "content": "\n".join(lines)},
        ],
        temperature=0.0,
    )
    content = _strip_json_fences(response.choices[0].message.content.strip())
    corrections = json.loads(content)
    if isinstance(corrections, dict):
        corrections = corrections.get("corrections", next(iter(corrections.values()), []))

    out = []
    for item in corrections:
        idx = item.get("id")
        if idx is None:
            continue
        seg = next((m for m in merged if m.id == idx), None)
        if seg is None:
            continue
        orig = (item.get("original") or "").strip()
        corr = (item.get("corrected") or "").strip()
        if orig and corr and orig != corr:
            out.append({
                "구분": "음성 대본",
                "시간": format_timestamp(seg.start),
                "수정 전": _red_to_html(orig),
                "수정 후": _red_to_html(corr),
                "교정 사유": item.get("reason", ""),
            })
    return out


# ─────────────────────────────────────────────
# 화면 프레임 처리
# ─────────────────────────────────────────────

# 전역 품질 설정 (속도 모드에서 app 쪽에서 오버라이드 가능)
_JPEG_QUALITY = 90           # 기존 95 → 90으로 낮춰도 OCR 정확도 영향 미미
_OCR_MAX_SIDE = 1920         # FHD 유지 (작은 글자 위해 너무 줄이면 안 됨)


def set_encode_quality(jpeg_quality: int = 90, max_side: int = 1920):
    """
    속도 조정용 전역 설정.
    - jpeg_quality: 낮을수록 업로드 용량↓ (권장 80~95)
    - max_side   : 이미지 최대 변. 1920 미만으로 낮추면 작은 글자 판독 저하 주의.
    """
    global _JPEG_QUALITY, _OCR_MAX_SIDE
    _JPEG_QUALITY = max(60, min(100, int(jpeg_quality)))
    _OCR_MAX_SIDE = max(720, int(max_side))


def _preprocess_for_ocr(image):
    """OCR 전 이미지 전처리: 리사이즈 + CLAHE + 언샤프."""
    h, w = image.shape[:2]
    if max(w, h) > _OCR_MAX_SIDE:
        scale = _OCR_MAX_SIDE / max(w, h)
        image = cv2.resize(image, (int(w * scale), int(h * scale)),
                           interpolation=cv2.INTER_LANCZOS4)

    lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
    l = clahe.apply(l)
    image = cv2.cvtColor(cv2.merge([l, a, b]), cv2.COLOR_LAB2BGR)

    gaussian = cv2.GaussianBlur(image, (0, 0), 2.0)
    image = cv2.addWeighted(image, 1.5, gaussian, -0.5, 0)

    return image


def encode_image(image):
    """전처리된 이미지를 JPEG Base64로 인코딩합니다."""
    image = _preprocess_for_ocr(image)
    _, buffer = cv2.imencode('.jpg', image, [int(cv2.IMWRITE_JPEG_QUALITY), _JPEG_QUALITY])
    return base64.b64encode(buffer).decode('utf-8')


def _phash(image_bgr, hash_size: int = 16) -> int:
    """
    간단한 perceptual hash (DCT 기반).
    자막 등 세밀한 변화까지 감지하도록 hash_size를 16으로 상향.
    """
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    resized = cv2.resize(gray, (hash_size * 4, hash_size * 4),
                         interpolation=cv2.INTER_AREA).astype(np.float32)
    dct = cv2.dct(resized)
    dct_low = dct[:hash_size, :hash_size]
    # DC 성분(좌상단) 제외 평균
    med = np.median(dct_low.flatten()[1:])
    bits = (dct_low > med).flatten()
    h = 0
    for b in bits:
        h = (h << 1) | int(b)
    return h


def _hamming(a: int, b: int) -> int:
    return bin(a ^ b).count("1")


def extract_and_filter_frames(
    video_path,
    sample_rate: float = 1.0,
    diff_threshold: float = 15.0,
    phash_threshold: int = 10,
):
    """
    영상에서 sample_rate 간격으로 프레임을 추출하고
    (1) 저해상도 평균차이로 빠른 필터링
    (2) pHash 해밍거리로 정밀 중복 제거

    ★ 속도 업그레이드:
    - `cap.grab()` + 조건부 `retrieve()` 패턴: 샘플링 간격에 해당하지 않는
      프레임은 디코딩 자체를 건너뛴다 → 대부분의 영상에서 2~3배 빠름.
    - 전체 프레임 수를 미리 계산해 진행 예측 가능.

    phash_threshold: 해밍거리 이하는 중복으로 간주 (기본 10 / 256비트 기준)
    """
    cap = cv2.VideoCapture(video_path)
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    frame_interval = max(1, int(fps * sample_rate))

    unique_frames = []
    prev_gray = None
    prev_hashes: List[int] = []
    frame_count = 0

    while True:
        # grab(): 디코딩 없이 다음 프레임으로 커서 이동 (매우 빠름)
        if not cap.grab():
            break

        # 샘플링 포인트에서만 실제 디코딩
        if frame_count % frame_interval == 0:
            ret, frame = cap.retrieve()
            if not ret:
                frame_count += 1
                continue

            small = cv2.resize(frame, (320, 180))
            gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)

            is_unique = True
            if prev_gray is not None:
                diff = cv2.absdiff(gray, prev_gray)
                if np.mean(diff) < diff_threshold:
                    # 1차 필터 탈락 → pHash로 정밀 비교
                    h = _phash(frame)
                    if any(_hamming(h, ph) <= phash_threshold for ph in prev_hashes[-5:]):
                        is_unique = False
                    else:
                        prev_hashes.append(h)
                else:
                    prev_hashes.append(_phash(frame))
            else:
                prev_hashes.append(_phash(frame))

            if is_unique:
                current_sec = frame_count / fps
                unique_frames.append({
                    "time": current_sec,
                    "time_str": format_timestamp(current_sec),
                    "base64": encode_image(frame),
                })
                prev_gray = gray

        frame_count += 1

    cap.release()
    return unique_frames


# ─────────────────────────────────────────────
# 화면 텍스트 맞춤법 검사
# ─────────────────────────────────────────────

_OCR_SPELL_PROMPT = f"""당신은 초정밀 한국어 OCR 및 맞춤법 교정 전문가입니다.
제공된 비디오 캡처 이미지들을 면밀히 분석하세요.

【작업 순서】
1. 각 이미지에서 보이는 **모든 한국어 텍스트**를 빠짐없이 `transcription`에 기록하세요.
   - 자막, 제목, 자료 화면, UI 라벨, 워터마크 등 위치 무관 전부 포함
   - 글자가 작거나 흐릿해도 최선을 다해 판독
2. 추출한 텍스트에서 맞춤법·오타 오류를 찾아 `corrections`에 담으세요.
3. **중복 제거**: 이전 이미지들과 **완전히 동일한 자막/텍스트**만 보인다면
   해당 이미지는 `corrections: []`로 비워두세요.

{_KO_SPELL_RULES}
{_SPELL_FEWSHOT}

【교정 규칙】
- 문장 전체를 반환하되, 오류 부분만 <red>단어</red>로 감싸세요.
- 확실한 오류만 교정 (불확실 → 원문 유지).
- OCR 판독 불확실한 글자는 `transcription`에는 추측, `corrections`에는 넣지 마세요.
- 이미지에 한국어 텍스트가 없거나 오류 없으면 `corrections: []`.
"""

_OCR_SCHEMA = {
    "name": "ocr_spell_results",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "results": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "id": {"type": "integer"},
                        "transcription": {"type": "string"},
                        "corrections": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "original":  {"type": "string"},
                                    "corrected": {"type": "string"},
                                    "reason":    {"type": "string"},
                                },
                                "required": ["original", "corrected", "reason"],
                            },
                        },
                    },
                    "required": ["id", "transcription", "corrections"],
                },
            }
        },
        "required": ["results"],
    },
}


def spell_check_frames(frames, api_key, batch_size=2, model="gpt-5.4", max_workers: int = 4):
    """
    프레임 이미지를 GPT Vision으로 OCR + 맞춤법 검사합니다.
    정확도 우선 모드: batch_size 기본값 2 (모델 집중도 유지).

    ★ 속도 업그레이드:
    - 배치 여러 개를 ThreadPoolExecutor로 병렬 호출. Vision API는 처리 시간이
      일반 텍스트보다 길어서 병렬화 효과가 크다. 권장 max_workers=3~5.
    """
    if not frames:
        return []

    client = OpenAI(api_key=api_key)

    # 배치 미리 생성
    batches: List[List[dict]] = [
        frames[i:i + batch_size] for i in range(0, len(frames), batch_size)
    ]

    def _process_batch(batch_idx: int, batch: List[dict]) -> List[dict]:
        content_items = [{"type": "text", "text": _OCR_SPELL_PROMPT}]
        for idx, f in enumerate(batch):
            content_items.append({"type": "text", "text": f"[{idx}번 이미지 — {f['time_str']}]"})
            content_items.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/jpeg;base64,{f['base64']}",
                    "detail": "high",
                },
            })

        def _call():
            return client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": content_items}],
                temperature=0.0,
                response_format={"type": "json_schema", "json_schema": _OCR_SCHEMA},
            )

        out: List[dict] = []
        try:
            response = call_with_retry(_call)
            parsed = json.loads(response.choices[0].message.content)
            for item in parsed.get("results", []):
                f_idx = item.get("id")
                if f_idx is None or not (0 <= f_idx < len(batch)):
                    continue
                for corr_item in item.get("corrections", []):
                    orig = (corr_item.get("original") or "").strip()
                    corr = (corr_item.get("corrected") or "").strip()
                    if orig and corr and orig != corr:
                        out.append({
                            "구분": "화면 텍스트",
                            "시간": batch[f_idx]["time_str"],
                            "수정 전": _red_to_html(orig),
                            "수정 후": _red_to_html(corr),
                            "교정 사유": corr_item.get("reason") or "",
                            "image_b64": batch[f_idx]["base64"],
                        })
        except Exception as e:
            print(f"화면 맞춤법 검사 오류 (배치 {batch_idx + 1}): {e}")
            try:
                out.extend(_fallback_text_mode_ocr(client, model, batch, batch_idx))
            except Exception as e2:
                print(f"  폴백도 실패: {e2}")
        return out

    # 병렬 실행 (시간순 보존)
    results: List[dict] = []
    if max_workers <= 1 or len(batches) == 1:
        for i, b in enumerate(batches):
            results.extend(_process_batch(i, b))
    else:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {ex.submit(_process_batch, i, b): i for i, b in enumerate(batches)}
            indexed: Dict[int, List[dict]] = {}
            for fut in as_completed(futures):
                i = futures[fut]
                try:
                    indexed[i] = fut.result()
                except Exception as e:
                    print(f"병렬 배치 {i + 1} 실패: {e}")
                    indexed[i] = []
            for i in sorted(indexed.keys()):
                results.extend(indexed[i])

    # 중복 교정 제거 (동일 오류가 여러 프레임에서 반복)
    seen = set()
    deduped = []
    for r in results:
        key = (
            re.sub(r"<[^>]+>", "", r["수정 전"]).strip(),
            re.sub(r"<[^>]+>", "", r["수정 후"]).strip(),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)

    return deduped


def _fallback_text_mode_ocr(client, model, batch, batch_idx):
    """OCR 텍스트 모드 폴백."""
    content_items = [{"type": "text", "text": _OCR_SPELL_PROMPT + (
        "\n응답은 순수 JSON 배열만 반환:\n"
        '[{"id":0,"transcription":"…","corrections":[{"original":"…","corrected":"…","reason":"…"}]}]'
    )}]
    for idx, f in enumerate(batch):
        content_items.append({"type": "text", "text": f"[{idx}번 이미지 — {f['time_str']}]"})
        content_items.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{f['base64']}", "detail": "high"},
        })

    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": content_items}],
        temperature=0.0,
    )
    content = _strip_json_fences(response.choices[0].message.content.strip())
    parsed = json.loads(content)
    if isinstance(parsed, dict):
        parsed = parsed.get("results", next(iter(parsed.values()), []))

    out = []
    for item in parsed:
        f_idx = item.get("id")
        if f_idx is None or not (0 <= f_idx < len(batch)):
            continue
        for corr_item in item.get("corrections", []):
            orig = (corr_item.get("original") or "").strip()
            corr = (corr_item.get("corrected") or "").strip()
            if orig and corr and orig != corr:
                out.append({
                    "구분": "화면 텍스트",
                    "시간": batch[f_idx]["time_str"],
                    "수정 전": _red_to_html(orig),
                    "수정 후": _red_to_html(corr),
                    "교정 사유": corr_item.get("reason", ""),
                    "image_b64": batch[f_idx]["base64"],
                })
    return out


# ─────────────────────────────────────────────
# 타임스탬프 → 프레임 추출
# ─────────────────────────────────────────────

def extract_frame_at_timestamp(video_path, time_str):
    """단일 타임스탬프의 프레임을 Base64로 반환."""
    try:
        ts = time_str.strip("[]")
        parts = list(map(int, ts.split(":")))
        if len(parts) == 2:
            seconds = parts[0] * 60 + parts[1]
        else:
            seconds = parts[0] * 3600 + parts[1] * 60 + parts[2]

        cap = cv2.VideoCapture(video_path)
        fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
        cap.set(cv2.CAP_PROP_POS_FRAMES, int(seconds * fps))
        ret, frame = cap.read()
        cap.release()
        if ret:
            return encode_image(frame)
    except Exception as e:
        print(f"프레임 추출 오류 ({time_str}): {e}")
    return ""


def attach_frames_to_audio_results(audio_results, video_path):
    """
    음성 교정 결과에 영상 프레임을 image_b64로 추가합니다.
    (VideoCapture를 한 번만 열고 여러 타임스탬프를 읽어 I/O 최소화)
    """
    if not audio_results:
        return audio_results

    cap = cv2.VideoCapture(video_path)
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0

    def _ts_to_sec(time_str: str) -> Optional[float]:
        try:
            parts = list(map(int, time_str.strip("[]").split(":")))
            if len(parts) == 2:
                return parts[0] * 60 + parts[1]
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
        except Exception:
            return None

    for r in audio_results:
        if r.get("image_b64"):
            continue
        seconds = _ts_to_sec(r.get("시간", ""))
        if seconds is None:
            continue
        try:
            cap.set(cv2.CAP_PROP_POS_FRAMES, int(seconds * fps))
            ret, frame = cap.read()
            if ret:
                r["image_b64"] = encode_image(frame)
            else:
                r["image_b64"] = ""
        except Exception as e:
            print(f"프레임 추출 오류 ({r.get('시간')}): {e}")
            r["image_b64"] = ""

    cap.release()
    return audio_results


# ─────────────────────────────────────────────
# 음성·화면 파이프라인 병렬 실행 (속도 업그레이드)
# ─────────────────────────────────────────────

def run_pipeline_parallel(
    video_path: str,
    audio_path: str,
    api_key: str,
    *,
    check_audio: bool = True,
    check_screen: bool = True,
    # 화면 파이프라인 파라미터
    sample_rate: float = 1.5,
    diff_threshold: float = 25.0,
    screen_batch_size: int = 2,
    screen_max_workers: int = 4,
    # 음성 파이프라인 파라미터
    model: str = "gpt-5.4",
    domain_hint: str = "",
    context_window: int = 2,
    use_sentence_merge: bool = True,
    audio_batch_size: int = 40,
    audio_max_workers: int = 3,
    stt_chunk_workers: int = 3,
    # ★ 캐싱 (안정성 업그레이드)
    stt_cache_dir: Optional[str] = None,
) -> Dict[str, Any]:
    """
    음성·화면 두 파이프라인을 **동시에** 실행합니다.

    ★ 안정성 옵션:
    - stt_cache_dir: 지정 시 Whisper 전사 결과를 파일 해시 기반으로 캐싱.
                     같은 영상·힌트 조합이면 즉시 재사용.

    Returns
    -------
    dict : {"audio": [...], "screen": [...], "errors": {...}, "stt_cached": bool}
    """

    errors: Dict[str, str] = {}
    stt_cached = False

    def _screen_pipeline() -> List[dict]:
        if not check_screen:
            return []
        try:
            frames = extract_and_filter_frames(video_path, sample_rate, diff_threshold)
            if not frames:
                return []
            return spell_check_frames(
                frames, api_key,
                batch_size=screen_batch_size,
                model=model,
                max_workers=screen_max_workers,
            )
        except Exception as e:
            errors["screen"] = str(e)
            return []

    def _audio_pipeline() -> List[dict]:
        nonlocal stt_cached
        if not check_audio:
            return []
        try:
            segments = None

            # 1) 캐시 조회
            cache_key = None
            if stt_cache_dir:
                try:
                    cache_key = compute_stt_cache_key(video_path, domain_hint)
                    cached = load_stt_cache(stt_cache_dir, cache_key)
                    if cached:
                        segments = cached
                        stt_cached = True
                        print(f"⚡ STT 캐시 적중 — 전사 단계 건너뜀 ({len(cached)} segs)")
                except Exception as e:
                    print(f"STT 캐시 조회 실패 (무시하고 계속): {e}")

            # 2) 캐시 미스 → 정상 전사
            if segments is None:
                if not extract_audio(video_path, audio_path):
                    errors["audio"] = "오디오 추출 실패"
                    return []
                segments = transcribe_audio(
                    audio_path, api_key,
                    domain_hint=domain_hint,
                    max_workers=stt_chunk_workers,
                )
                if stt_cache_dir and cache_key and segments:
                    save_stt_cache(stt_cache_dir, cache_key, segments)

            if not segments:
                return []

            results = spell_check_segments(
                segments, api_key,
                context_window=context_window,
                model=model,
                use_sentence_merge=use_sentence_merge,
                batch_size=audio_batch_size,
                max_workers=audio_max_workers,
            )
            return attach_frames_to_audio_results(results, video_path)
        except Exception as e:
            errors["audio"] = str(e)
            return []

    # 두 파이프라인 병렬 실행
    with ThreadPoolExecutor(max_workers=2) as ex:
        fut_screen = ex.submit(_screen_pipeline)
        fut_audio  = ex.submit(_audio_pipeline)
        screen_results = fut_screen.result()
        audio_results  = fut_audio.result()

    return {
        "audio":  audio_results,
        "screen": screen_results,
        "errors": errors,
        "stt_cached": stt_cached,
    }


# ─────────────────────────────────────────────
# 내부 헬퍼
# ─────────────────────────────────────────────

def _strip_json_fences(text):
    """마크다운 코드 펜스(```json ... ```)를 제거합니다."""
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()


def _red_to_html(text):
    """<red>단어</red> 태그를 인라인 HTML 강조 스타일로 변환합니다."""
    return (text
            .replace("<red>", "<span style='color:red; font-weight:bold;'>")
            .replace("</red>", "</span>"))
