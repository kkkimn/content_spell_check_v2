import os
import re
import io
import base64
import datetime
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import pandas as pd
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from core_video import (
    extract_audio,
    transcribe_audio,
    spell_check_segments,
    extract_and_filter_frames,
    spell_check_frames,
    attach_frames_to_audio_results,
    run_pipeline_parallel,
    set_encode_quality,
)

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(page_title="영상 맞춤법 검사기", page_icon="🎥", layout="wide")

st.title("🎥 영상 대본 및 화면 글자 자동 교정기")
st.markdown(
    "MP4 영상을 업로드하면 AI가 **음성 대본(STT)** 및 **화면 속 텍스트(OCR)**를 분석하여 "
    "맞춤법·오타·띄어쓰기 오류를 잡아줍니다."
)

# ─────────────────────────────────────────────
# 공통 CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
.card {
    background: #fff;
    border: 1px solid #e0e4ec;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 14px;
    font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.card-header {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 12px;
    font-size: 13px;
    color: #666;
    border-bottom: 1px solid #f0f0f0;
    padding-bottom: 8px;
}
.badge-audio  { background:#4f8ef7; color:white; border-radius:5px; padding:2px 9px; font-size:12px; font-weight:bold; }
.badge-screen { background:#22b07d; color:white; border-radius:5px; padding:2px 9px; font-size:12px; font-weight:bold; }
.timestamp    { background:#f3f4f6; border-radius:4px; padding:2px 8px; font-size:12px; color:#555; font-family:monospace; }
.label-before { font-size:12px; font-weight:bold; color:#999; margin-bottom:4px; }
.label-after  { font-size:12px; font-weight:bold; color:#4f8ef7; margin-bottom:4px; }
.text-before  { background:#fff8f8; border-left:3px solid #f87171; border-radius:4px; padding:8px 12px; font-size:15px; color:#333; line-height:1.7; margin-bottom:8px; }
.text-after   { background:#f0fdf4; border-left:3px solid #34d399; border-radius:4px; padding:8px 12px; font-size:15px; color:#333; line-height:1.7; margin-bottom:6px; }
.reason-box   { font-size:12px; color:#888; margin-top:4px; }
.no-result    { text-align:center; padding:40px; color:#aaa; font-size:15px; }
.info-box     { background:#f8faff; border:1px solid #d0dcf5; border-radius:10px; padding:18px 22px; margin-bottom:20px; }
[data-testid="stSidebar"] hr { margin-top: 0.4rem !important; margin-bottom: 0.4rem !important; }
.info-box h4  { margin:0 0 12px 0; color:#3a5bbf; font-size:15px; }
.info-grid    { display:grid; grid-template-columns:1fr 1fr 1fr; gap:12px; }
.info-item label { font-size:11px; color:#888; display:block; margin-bottom:3px; }
.info-item span  { font-size:14px; font-weight:bold; color:#222; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 엑셀 생성 함수
# ─────────────────────────────────────────────
def _b64_to_xl_image(b64_str, thumb_w=240, thumb_h=135):
    """Base64 JPEG 문자열을 openpyxl용 XLImage 객체로 변환합니다."""
    raw = base64.b64decode(b64_str)
    pil = PILImage.open(io.BytesIO(raw)).convert("RGB")
    pil.thumbnail((thumb_w, thumb_h), PILImage.LANCZOS)
    buf = io.BytesIO()
    pil.save(buf, format="JPEG", quality=85)
    buf.seek(0)
    return XLImage(buf)


def make_excel(audio_results, screen_results, reviewer, lesson_name, lesson_num, review_date=None):
    wb = Workbook()

    def strip_html(text):
        return re.sub(r'<[^>]+>', '', text or '')

    # ── 공통 스타일 ───────────────────────────────
    header_fill  = PatternFill("solid", start_color="1E2533")
    audio_fill   = PatternFill("solid", start_color="EBF3FF")
    screen_fill  = PatternFill("solid", start_color="EAFAF3")
    all_fill     = PatternFill("solid", start_color="F5F7FA")
    meta_fill    = PatternFill("solid", start_color="F0F4FF")
    white_fill   = PatternFill("solid", start_color="FFFFFF")

    header_font  = Font(name="맑은 고딕", bold=True, color="FFFFFF",  size=11)
    title_font   = Font(name="맑은 고딕", bold=True, color="1E2533",  size=16)
    body_font    = Font(name="맑은 고딕",             color="111111",  size=10)
    meta_label_f = Font(name="맑은 고딕", bold=True,  color="888888",  size=9)
    meta_value_f = Font(name="맑은 고딕", bold=True,  color="1E2533",  size=11)

    thin   = Side(style="thin",   color="CCCCCC")
    medium = Side(style="medium", color="AAAAAA")
    thin_b   = Border(left=thin,   right=thin,   top=thin,   bottom=thin)
    thick_b  = Border(left=medium, right=medium, top=medium, bottom=medium)
    center   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_top = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

    IMG_ROW_H = 105   # 이미지 행 높이 (포인트, ≈ 140px)
    IMG_COL   = 6     # 이미지 컬럼 번호 (F) — 조치결과(G) 앞
    IMG_COL_W = 36    # 이미지 컬럼 너비 (문자 단위)
    action_fill = PatternFill("solid", start_color="FFFDE7")  # 연노랑 — 사용자 입력 유도

    def write_meta(ws, has_image_col=False):
        """제목 + 메타(검토자/차시명/차시) 블록을 씁니다."""
        last_col = "G" if has_image_col else "F"

        ws.merge_cells(f"A1:{last_col}1")
        ws["A1"] = "영상 맞춤법 교정 결과 보고서"
        ws["A1"].font      = title_font
        ws["A1"].fill      = white_fill
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 36

        ws.row_dimensions[2].height = 6
        date_str = review_date.strftime("%Y년 %m월 %d일") if review_date else "-"
        for i, (lbl, val) in enumerate([
            ("검토자",  reviewer    or "-"),
            ("차시명",  lesson_name or "-"),
            ("차시",    lesson_num  or "-"),
            ("검토일자", date_str),
        ]):
            r = 3 + i
            ws.merge_cells(f"A{r}:B{r}")
            ws.merge_cells(f"C{r}:{last_col}{r}")
            ws[f"A{r}"] = lbl;  ws[f"A{r}"].font = meta_label_f
            ws[f"A{r}"].fill = meta_fill; ws[f"A{r}"].alignment = center
            ws[f"A{r}"].border = thin_b
            ws[f"C{r}"] = val;  ws[f"C{r}"].font = meta_value_f
            ws[f"C{r}"].fill = meta_fill; ws[f"C{r}"].alignment = left_top
            ws[f"C{r}"].border = thin_b
            ws.row_dimensions[r].height = 22
        ws.row_dimensions[7].height = 8  # 여백 (메타 4개 → 3~6행, 여백 7행)

    def write_headers(ws, has_image_col=False):
        """컬럼 헤더 행(8번)을 씁니다."""
        if has_image_col:
            headers = ["구분", "시간", "수정 전", "수정 후", "교정 사유", "수정 전 화면 이미지", "조치결과", "수정 후 화면 이미지"]
        else:
            headers = ["구분", "시간", "수정 전", "수정 후", "교정 사유", "조치결과"]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=8, column=col, value=h)
            c.font = header_font; c.fill = header_fill
            c.alignment = center; c.border = thick_b
        ws.row_dimensions[8].height = 24

    def write_rows(ws, results, badge_label, row_fill, has_image_col=False):
        """데이터 행을 쓰고 이미지가 있으면 삽입합니다."""
        total_cols = 8 if has_image_col else 6
        if not results:
            span = f"A9:{get_column_letter(total_cols)}9"
            ws.merge_cells(span)
            ws["A9"] = "교정 항목 없음"
            ws["A9"].alignment = center; ws["A9"].font = body_font
            ws.row_dimensions[9].height = 22
            return

        for i, r in enumerate(results):
            row_idx = 9 + i
            gbn  = badge_label if badge_label else r.get("구분", "")
            fill = row_fill if i % 2 == 0 else white_fill
            vals = [
                gbn,
                r.get("시간", ""),
                strip_html(r.get("수정 전", "")),
                strip_html(r.get("수정 후", "")),
                r.get("교정 사유", ""),
            ]
            for col, v in enumerate(vals, 1):
                c = ws.cell(row=row_idx, column=col, value=v)
                c.font = body_font; c.fill = fill; c.border = thin_b
                c.alignment = center if col <= 2 else left_top

            # 이미지 삽입 (col 6)
            # 조치결과 칸은 이미지 뒤 col 7에 위치
            if has_image_col:
                b64 = r.get("image_b64", "")
                img_cell = ws.cell(row=row_idx, column=IMG_COL)
                img_cell.fill = fill; img_cell.border = thin_b
                if b64:
                    try:
                        xl_img = _b64_to_xl_image(b64)
                        xl_img.anchor = f"{get_column_letter(IMG_COL)}{row_idx}"
                        ws.add_image(xl_img)
                    except Exception as e:
                        img_cell.value = "이미지 오류"
                ws.row_dimensions[row_idx].height = IMG_ROW_H
                # 조치결과 칸 (col 7) — 연노랑 빈칸
                ac = ws.cell(row=row_idx, column=7, value="")
                ac.fill = action_fill; ac.border = thin_b; ac.alignment = center
                # 수정 후 화면 이미지 칸 (col 8) — 연노랑 빈칸, 사용자가 직접 삽입
                af = ws.cell(row=row_idx, column=8, value="")
                af.fill = action_fill; af.border = thin_b; af.alignment = center
            else:
                ws.row_dimensions[row_idx].height = 40

    def set_col_widths(ws, has_image_col=False):
        # 구분, 시간, 수정전, 수정후, 교정사유[, 수정전이미지, 조치결과, 수정후이미지]
        widths = [12, 12, 34, 34, 28]
        if has_image_col:
            widths += [IMG_COL_W, 9, IMG_COL_W]
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

    # ── 시트 1: 음성 대본 (이미지 포함) ─────────────
    ws_audio = wb.active
    ws_audio.title = "음성 대본"
    write_meta(ws_audio, has_image_col=True)
    write_headers(ws_audio, has_image_col=True)
    write_rows(ws_audio, audio_results, "음성 대본", audio_fill, has_image_col=True)
    set_col_widths(ws_audio, has_image_col=True)

    # ── 시트 2: 화면 텍스트 (이미지 포함) ───────────
    ws_screen = wb.create_sheet("화면 텍스트")
    write_meta(ws_screen, has_image_col=True)
    write_headers(ws_screen, has_image_col=True)
    write_rows(ws_screen, screen_results, "화면 텍스트", screen_fill, has_image_col=True)
    set_col_widths(ws_screen, has_image_col=True)

    # ── 시트 3: 통합 (시간순, 이미지 포함) ──────────
    all_results = sorted(
        list(audio_results) + list(screen_results),
        key=lambda x: x.get("시간", "")
    )
    ws_all = wb.create_sheet("통합")
    write_meta(ws_all, has_image_col=True)
    write_headers(ws_all, has_image_col=True)
    write_rows(ws_all, all_results, "", all_fill, has_image_col=True)
    set_col_widths(ws_all, has_image_col=True)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def make_excel_combined(all_video_results, reviewer, lesson_name, lesson_num, review_date=None):
    """
    여러 영상의 결과를 하나의 엑셀 파일로 통합합니다.
    구조: [영상명_음성], [영상명_화면], ... + [전체통합] 시트
    """
    wb = Workbook()
    first = True  # 첫 번째 시트는 active 시트 사용

    def strip_html(text):
        return re.sub(r'<[^>]+>', '', text or '')

    # 공통 스타일
    header_fill  = PatternFill("solid", start_color="1E2533")
    audio_fill   = PatternFill("solid", start_color="EBF3FF")
    screen_fill  = PatternFill("solid", start_color="EAFAF3")
    all_fill     = PatternFill("solid", start_color="F5F7FA")
    meta_fill    = PatternFill("solid", start_color="F0F4FF")
    white_fill   = PatternFill("solid", start_color="FFFFFF")
    action_fill  = PatternFill("solid", start_color="FFFDE7")

    header_font  = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)
    title_font   = Font(name="맑은 고딕", bold=True, color="1E2533", size=16)
    body_font    = Font(name="맑은 고딕", color="111111", size=10)
    meta_label_f = Font(name="맑은 고딕", bold=True, color="888888", size=9)
    meta_value_f = Font(name="맑은 고딕", bold=True, color="1E2533", size=11)
    video_title_font = Font(name="맑은 고딕", bold=True, color="3A5BBF", size=13)

    thin   = Side(style="thin",   color="CCCCCC")
    medium = Side(style="medium", color="AAAAAA")
    thin_b  = Border(left=thin,   right=thin,   top=thin,   bottom=thin)
    thick_b = Border(left=medium, right=medium, top=medium, bottom=medium)
    center   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_top = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

    IMG_ROW_H = 105
    IMG_COL   = 6
    IMG_COL_W = 36

    def _write_sheet(ws, results, badge_label, row_fill, video_name, has_img=True):
        """한 시트에 메타 + 헤더 + 데이터를 씁니다."""
        last_col = "H" if has_img else "F"
        date_str = review_date.strftime("%Y년 %m월 %d일") if review_date else "-"

        # 제목
        ws.merge_cells(f"A1:{last_col}1")
        ws["A1"] = "영상 맞춤법 교정 결과 보고서"
        ws["A1"].font = title_font; ws["A1"].fill = white_fill
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 36

        # 영상명 표시 (통합본에서 어느 영상인지 구분)
        ws.merge_cells(f"A2:{last_col}2")
        ws["A2"] = f"📹 {video_name}"
        ws["A2"].font = video_title_font
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[2].height = 22

        # 메타 (행 3~6)
        ws.row_dimensions[3].height = 4  # 여백
        for i, (lbl, val) in enumerate([
            ("검토자",  reviewer    or "-"),
            ("차시명",  lesson_name or "-"),
            ("차시",    lesson_num  or "-"),
            ("검토일자", date_str),
        ]):
            r = 4 + i
            ws.merge_cells(f"A{r}:B{r}")
            ws.merge_cells(f"C{r}:{last_col}{r}")
            ws[f"A{r}"] = lbl; ws[f"A{r}"].font = meta_label_f
            ws[f"A{r}"].fill = meta_fill; ws[f"A{r}"].alignment = center
            ws[f"A{r}"].border = thin_b
            ws[f"C{r}"] = val; ws[f"C{r}"].font = meta_value_f
            ws[f"C{r}"].fill = meta_fill; ws[f"C{r}"].alignment = left_top
            ws[f"C{r}"].border = thin_b
            ws.row_dimensions[r].height = 22
        ws.row_dimensions[8].height = 6  # 여백

        # 헤더 (행 9)
        if has_img:
            headers = ["구분", "시간", "수정 전", "수정 후", "교정 사유",
                       "수정 전 화면 이미지", "조치결과", "수정 후 화면 이미지"]
        else:
            headers = ["구분", "시간", "수정 전", "수정 후", "교정 사유", "조치결과"]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=9, column=col, value=h)
            c.font = header_font; c.fill = header_fill
            c.alignment = center; c.border = thick_b
        ws.row_dimensions[9].height = 24

        # 데이터 (행 10~)
        total_cols = 8 if has_img else 6
        if not results:
            span = f"A10:{get_column_letter(total_cols)}10"
            ws.merge_cells(span)
            ws["A10"] = "교정 항목 없음"
            ws["A10"].alignment = center; ws["A10"].font = body_font
            ws.row_dimensions[10].height = 22
        else:
            for i, r in enumerate(results):
                row_idx = 10 + i
                gbn  = badge_label if badge_label else r.get("구분", "")
                fill = row_fill if i % 2 == 0 else white_fill
                vals = [gbn, r.get("시간",""), strip_html(r.get("수정 전","")),
                        strip_html(r.get("수정 후","")), r.get("교정 사유","")]
                for col, v in enumerate(vals, 1):
                    c = ws.cell(row=row_idx, column=col, value=v)
                    c.font = body_font; c.fill = fill; c.border = thin_b
                    c.alignment = center if col <= 2 else left_top
                if has_img:
                    b64 = r.get("image_b64","")
                    img_cell = ws.cell(row=row_idx, column=IMG_COL)
                    img_cell.fill = fill; img_cell.border = thin_b
                    if b64:
                        try:
                            xl_img = _b64_to_xl_image(b64)
                            xl_img.anchor = f"{get_column_letter(IMG_COL)}{row_idx}"
                            ws.add_image(xl_img)
                        except Exception:
                            img_cell.value = "이미지 오류"
                    ws.row_dimensions[row_idx].height = IMG_ROW_H
                    ac = ws.cell(row=row_idx, column=7, value="")
                    ac.fill = action_fill; ac.border = thin_b; ac.alignment = center
                    af = ws.cell(row=row_idx, column=8, value="")
                    af.fill = action_fill; af.border = thin_b; af.alignment = center
                else:
                    ws.row_dimensions[row_idx].height = 40

        # 컬럼 너비
        widths = [12, 12, 34, 34, 28]
        if has_img:
            widths += [IMG_COL_W, 9, IMG_COL_W]
        else:
            widths += [18]
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

    # ── 영상별 시트 생성 ─────────────────────────────
    all_combined = []   # 전체통합 시트용
    for vr in all_video_results:
        vname = vr["name"]
        audio  = vr["audio"]
        screen = vr["screen"]

        # 시트명 최대 31자 제한 (엑셀 규칙)
        short = vname[:12] if len(vname) > 12 else vname

        # 음성 시트
        ws_a = wb.active if first else wb.create_sheet(f"{short}_음성")
        if first:
            ws_a.title = f"{short}_음성"
            first = False
        _write_sheet(ws_a, audio, "음성 대본", audio_fill, vname)

        # 화면 시트
        ws_s = wb.create_sheet(f"{short}_화면")
        _write_sheet(ws_s, screen, "화면 텍스트", screen_fill, vname)

        # 전체통합용 데이터에 영상명 추가
        for r in audio + screen:
            all_combined.append(dict(r, _video=vname))

    # ── 전체 통합 시트 ───────────────────────────────
    all_combined.sort(key=lambda x: (x.get("_video",""), x.get("시간","")))
    ws_all = wb.create_sheet("📋 전체통합")

    # 전체통합은 영상명 컬럼을 맨 앞에 추가
    last_col_all = "I"
    date_str_all = review_date.strftime("%Y년 %m월 %d일") if review_date else "-"
    ws_all.merge_cells(f"A1:{last_col_all}1")
    ws_all["A1"] = "영상 맞춤법 교정 결과 보고서 — 전체 통합"
    ws_all["A1"].font = title_font; ws_all["A1"].fill = white_fill
    ws_all["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws_all.row_dimensions[1].height = 36
    ws_all.row_dimensions[2].height = 4
    for i, (lbl, val) in enumerate([
        ("검토자",  reviewer    or "-"),
        ("차시명",  lesson_name or "-"),
        ("차시",    lesson_num  or "-"),
        ("검토일자", date_str_all),
    ]):
        r = 3 + i
        ws_all.merge_cells(f"A{r}:B{r}")
        ws_all.merge_cells(f"C{r}:{last_col_all}{r}")
        ws_all[f"A{r}"] = lbl; ws_all[f"A{r}"].font = meta_label_f
        ws_all[f"A{r}"].fill = meta_fill; ws_all[f"A{r}"].alignment = center
        ws_all[f"A{r}"].border = thin_b
        ws_all[f"C{r}"] = val; ws_all[f"C{r}"].font = meta_value_f
        ws_all[f"C{r}"].fill = meta_fill; ws_all[f"C{r}"].alignment = left_top
        ws_all[f"C{r}"].border = thin_b
        ws_all.row_dimensions[r].height = 22
    ws_all.row_dimensions[7].height = 6

    # 전체통합 헤더 (영상명 컬럼 추가)
    all_headers = ["영상명", "구분", "시간", "수정 전", "수정 후",
                   "교정 사유", "수정 전 화면 이미지", "조치결과", "수정 후 화면 이미지"]
    for col, h in enumerate(all_headers, 1):
        c = ws_all.cell(row=8, column=col, value=h)
        c.font = header_font; c.fill = header_fill
        c.alignment = center; c.border = thick_b
    ws_all.row_dimensions[8].height = 24

    if not all_combined:
        ws_all.merge_cells("A9:I9")
        ws_all["A9"] = "교정 항목 없음"
        ws_all["A9"].alignment = center; ws_all["A9"].font = body_font
        ws_all.row_dimensions[9].height = 22
    else:
        for i, r in enumerate(all_combined):
            row_idx = 9 + i
            fill = all_fill if i % 2 == 0 else white_fill
            vals = [
                r.get("_video",""),
                r.get("구분",""),
                r.get("시간",""),
                strip_html(r.get("수정 전","")),
                strip_html(r.get("수정 후","")),
                r.get("교정 사유",""),
            ]
            for col, v in enumerate(vals, 1):
                c = ws_all.cell(row=row_idx, column=col, value=v)
                c.font = body_font; c.fill = fill; c.border = thin_b
                c.alignment = center if col <= 3 else left_top
            # 이미지 (col 7)
            b64 = r.get("image_b64","")
            img_cell = ws_all.cell(row=row_idx, column=7)
            img_cell.fill = fill; img_cell.border = thin_b
            if b64:
                try:
                    xl_img = _b64_to_xl_image(b64)
                    xl_img.anchor = f"G{row_idx}"
                    ws_all.add_image(xl_img)
                except Exception:
                    img_cell.value = "이미지 오류"
            ws_all.row_dimensions[row_idx].height = IMG_ROW_H
            ac = ws_all.cell(row=row_idx, column=8, value="")
            ac.fill = action_fill; ac.border = thin_b; ac.alignment = center
            af = ws_all.cell(row=row_idx, column=9, value="")
            af.fill = action_fill; af.border = thin_b; af.alignment = center

    # 전체통합 컬럼 너비
    for col, w in enumerate([20, 12, 12, 34, 34, 28, IMG_COL_W, 9, IMG_COL_W], 1):
        ws_all.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# 결과 카드 렌더링
# ─────────────────────────────────────────────
def render_result_cards(results, badge_type="audio"):
    if not results:
        st.markdown('<div class="no-result">✅ 교정이 필요한 항목이 없습니다.</div>', unsafe_allow_html=True)
        return
    badge_class = "badge-audio" if badge_type == "audio" else "badge-screen"
    badge_label = "음성 대본" if badge_type == "audio" else "화면 텍스트"
    for i, r in enumerate(results, 1):
        reason_html = f'<div class="reason-box">📌 {r.get("교정 사유","")}</div>' if r.get("교정 사유") else ""
        st.markdown(f"""
        <div class="card">
            <div class="card-header">
                <span class="{badge_class}">{badge_label}</span>
                <span class="timestamp">{r.get("시간","")}</span>
                <span style="margin-left:auto;color:#bbb;font-size:12px;">#{i}</span>
            </div>
            <div class="label-before">수정 전</div>
            <div class="text-before">{r.get("수정 전","")}</div>
            <div class="label-after">수정 후</div>
            <div class="text-after">{r.get("수정 후","")}</div>
            {reason_html}
        </div>
        """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 사이드바 — API 키
# ─────────────────────────────────────────────
# st.sidebar.header("⚙️ 설정")
try:
    api_key = st.secrets["OPENAI_API_KEY"]
except Exception:
    api_key = ""

if not api_key:
    st.sidebar.warning("아래에 OpenAI API 키를 직접 입력하세요.")
    api_key = st.sidebar.text_input("OpenAI API Key", type="password", help="sk-... 형태의 키")
    if not api_key:
        st.info("👈 사이드바에서 OpenAI API 키를 입력하면 검사가 시작됩니다.")
        st.stop()

# ─────────────────────────────────────────────
# 상단 헤더 — Deploy 옆 로고
# ─────────────────────────────────────────────
_logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ARASoft로고.png")
if os.path.exists(_logo_path):
    st.markdown(
        """
        <style>
        .st-emotion-cache-4xtz07 {
            height: 3rem !important;
            margin-top: 2.25rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    try:
        st.logo(_logo_path, size="large")
    except TypeError:
        # size 파라미터 미지원 버전
        st.logo(_logo_path)
    except AttributeError:
        # st.logo 자체 미지원 구버전 fallback
        _logo_b64 = base64.b64encode(open(_logo_path, "rb").read()).decode()
        st.markdown(
            f"""
            <style>
            [data-testid="stToolbar"]::before {{
                content: "";
                display: inline-block;
                background-image: url("data:image/png;base64,{_logo_b64}");
                background-size: contain;
                background-repeat: no-repeat;
                background-position: center;
                width: 120px;
                height: 34px;
                vertical-align: middle;
                margin-right: 8px;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )

# ─────────────────────────────────────────────
# 사이드바 — 검사 대상
# ─────────────────────────────────────────────
st.sidebar.divider()
st.sidebar.header("🎯 검사 대상")
check_audio  = st.sidebar.checkbox("🎧 음성 대본 검사", value=True)
check_screen = st.sidebar.checkbox("🖼️ 화면 텍스트 검사", value=True)
if not check_audio and not check_screen:
    st.warning("👈 검사 대상을 최소 하나 이상 체크해주세요.")
    st.stop()

# ─────────────────────────────────────────────
# 사이드바 — 문서 정보 입력
# ─────────────────────────────────────────────
st.sidebar.divider()
st.sidebar.header("📄 문서 정보 (엑셀 헤더)")
reviewer    = st.sidebar.text_input("검토자",  placeholder="예) 홍길동")
lesson_name = st.sidebar.text_input("차시명",  placeholder="예) 1강 - 파이썬 기초")
lesson_num  = st.sidebar.text_input("차시",    placeholder="예) 1")
review_date = st.sidebar.date_input("검토일자", value=datetime.date.today())

# ─────────────────────────────────────────────
# 사이드바 — 고급 설정
# ─────────────────────────────────────────────
st.sidebar.divider()
st.sidebar.header("🔧 고급 설정")

selected_model = st.sidebar.radio(
    "🤖 AI 모델 선택",
    options=["gpt-5.4", "gpt-5.4-mini"],
    index=1,
    help="gpt-5.4: 최고 정확도 / gpt-5.4-mini: 빠르고 저렴 (약 6배 저렴, 94% 성능)"
)
model_label = "GPT-5.4" if selected_model == "gpt-5.4" else "GPT-5.4 mini"

with st.sidebar.expander("화면 프레임 설정", expanded=False):
    sample_rate    = st.slider("프레임 추출 간격 (초)", 0.5, 5.0, 1.5, 0.5)
    diff_threshold = st.slider("화면 변화 감지 임계값", 5.0, 50.0, 25.0, 5.0)
    # 정확도 우선: 기본 배치 2 (모델 집중도 ↑)
    batch_size     = st.selectbox("Vision API 배치 크기", [1, 2, 3, 4, 5], index=1,
                                   help="작을수록 정확도 ↑ (속도 ↓). 기본 2 권장.")

with st.sidebar.expander("음성 검사 설정", expanded=False):
    context_window     = st.slider("문맥 참고 세그먼트 수", 1, 5, 2, 1,
                                    help="교정 대상 앞뒤로 참고할 세그먼트 수")
    use_sentence_merge = st.checkbox("문장 단위 세그먼트 병합", value=True,
                                      help="Whisper가 짧게 끊어낸 세그먼트를 문장 단위로 병합 → "
                                           "문맥 오류(이중피동·어미 혼동) 검출률 향상")
    audio_batch_size   = st.slider("음성 교정 배치 크기", 10, 80, 40, 10,
                                    help="한 번에 검사할 문장 수. 작을수록 정확도 ↑")

with st.sidebar.expander("🎯 도메인 힌트 (정확도 향상)", expanded=False):
    domain_hint = st.text_area(
        "영상의 주요 용어·고유명사",
        placeholder="예) 파이썬 수업, 리스트 컴프리헨션, 재귀 함수, 객체지향",
        help="Whisper STT에 힌트로 전달되어 고유명사·전문용어 오인식을 크게 줄입니다.",
        height=80,
    )

with st.sidebar.expander("🚀 속도 설정", expanded=False):
    parallel_videos = st.slider(
        "영상 동시 처리 개수", 1, 4, 2, 1,
        help="여러 MP4를 동시에 처리합니다. OpenAI rate limit 고려 1~3 권장.",
    )
    parallel_pipelines = st.checkbox(
        "한 영상 내에서 음성·화면 검사 병렬 실행", value=True,
        help="독립적인 두 파이프라인을 동시에 실행 → 전체 시간이 긴 쪽 하나로 수렴",
    )
    audio_workers = st.slider(
        "음성 배치 병렬 호출 수", 1, 6, 3, 1,
        help="맞춤법 교정 배치를 병렬로 호출합니다. 너무 높이면 429 rate limit 발생 가능.",
    )
    screen_workers = st.slider(
        "화면 OCR 배치 병렬 호출 수", 1, 6, 4, 1,
        help="Vision API 배치 병렬 호출. Vision은 응답이 느려 병렬 효과 큼.",
    )
    jpeg_quality = st.slider(
        "JPEG 품질 (OCR 업로드)", 70, 100, 90, 5,
        help="낮추면 업로드가 빨라지지만 작은 글자 판독률↓. 90 권장.",
    )
    # 전역 인코딩 품질 적용
    set_encode_quality(jpeg_quality=jpeg_quality, max_side=1920)

with st.sidebar.expander("🛡️ 안정성 설정", expanded=False):
    use_stt_cache = st.checkbox(
        "STT 결과 캐싱 (같은 영상 재실행 시 즉시 결과)", value=True,
        help="영상 파일 내용 해시 + 도메인 힌트 조합을 키로 사용합니다. "
             "같은 영상을 다시 돌릴 때 Whisper 비용/시간 0.",
    )
    stt_chunk_workers = st.slider(
        "긴 오디오 청크 병렬 전사 수", 1, 4, 3, 1,
        help="25MB 초과 오디오가 자동으로 청크 분할될 때 병렬 전사 워커 수.",
    )
    if st.button("🗑️ STT 캐시 비우기", use_container_width=True):
        cache_root = os.path.join(tempfile.gettempdir(), "spellcheck_stt_cache")
        import shutil as _sh
        if os.path.isdir(cache_root):
            _sh.rmtree(cache_root, ignore_errors=True)
            st.success("STT 캐시 삭제 완료")
        else:
            st.info("캐시 폴더가 없습니다.")

# ─────────────────────────────────────────────
# 메인 — 문서 정보 미리보기
# ─────────────────────────────────────────────
review_date_str = review_date.strftime("%Y년 %m월 %d일") if review_date else "-"
if reviewer or lesson_name or lesson_num:
    st.markdown(f"""
    <div class="info-box">
        <h4>📄 엑셀 문서 정보</h4>
        <div class="info-grid" style="grid-template-columns:1fr 1fr 1fr 1fr;">
            <div class="info-item"><label>검토자</label><span>{reviewer or "-"}</span></div>
            <div class="info-item"><label>차시명</label><span>{lesson_name or "-"}</span></div>
            <div class="info-item"><label>차시</label><span>{lesson_num or "-"}</span></div>
            <div class="info-item"><label>검토일자</label><span>{review_date_str}</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 메인 — 파일 업로드 (다중 파일)
# ─────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "검사할 동영상 파일(MP4)을 업로드하세요. (여러 개 동시 업로드 가능)",
    type=["mp4"],
    accept_multiple_files=True
)

if uploaded_files:
    # 업로드된 영상 미리보기 (3열 그리드, 자동 줄바꿈)
    st.markdown(f"**총 {len(uploaded_files)}개 영상 업로드됨**")
    for row_start in range(0, len(uploaded_files), 3):
        row_files = uploaded_files[row_start:row_start+3]
        cols = st.columns(3)
        for col, f in zip(cols, row_files):
            with col:
                st.caption(f"📹 {f.name}")
                st.video(f)

    st.caption(f"🤖 사용 모델: **{model_label}**  ·  🚀 영상 동시 처리: **{parallel_videos}개**  ·  "
               f"파이프라인 병렬: **{'ON' if parallel_pipelines else 'OFF'}**  ·  "
               f"🛡️ STT 캐시: **{'ON' if use_stt_cache else 'OFF'}**")

    if st.button("🚀 맞춤법 검사 시작", type="primary", use_container_width=True):

        import time as _time
        _t0 = _time.time()

        # STT 캐시 디렉토리 (user cache temp 아래)
        stt_cache_dir = None
        if use_stt_cache:
            stt_cache_dir = os.path.join(tempfile.gettempdir(), "spellcheck_stt_cache")
            os.makedirs(stt_cache_dir, exist_ok=True)

        # ── 1) 업로드 파일을 모두 고유 임시 디렉토리에 저장 ───
        work_dir = tempfile.mkdtemp(prefix="spellcheck_")
        video_infos = []   # [{idx, name, filename, video_path, audio_path}]
        all_video_results = []   # try 바깥에서 미리 초기화 (예외 발생 시 NameError 방지)
        try:
            with st.spinner(f"파일 {len(uploaded_files)}개 준비 중..."):
                for file_idx, uploaded_file in enumerate(uploaded_files):
                    video_name = uploaded_file.name.rsplit(".", 1)[0]
                    video_path = os.path.join(work_dir, f"video_{file_idx}.mp4")
                    audio_path = os.path.join(work_dir, f"audio_{file_idx}.mp3")
                    with open(video_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    video_infos.append({
                        "idx": file_idx,
                        "name": video_name,
                        "filename": uploaded_file.name,
                        "video_path": video_path,
                        "audio_path": audio_path,
                    })

            # ── 2) 진행 상태 UI ─────────────────────────
            overall_prog = st.progress(0.0, text=f"0 / {len(video_infos)}개 영상 처리 완료")
            status_area = st.empty()
            status_dict = {vi["idx"]: "⏳ 대기 중" for vi in video_infos}

            def _render_status():
                lines = []
                for vi in video_infos:
                    lines.append(f"- **{vi['filename']}** — {status_dict[vi['idx']]}")
                status_area.markdown("\n".join(lines))
            _render_status()

            # ── 3) 단일 영상 처리 함수 (워커 스레드에서 실행) ───
            def _process_one(vi):
                try:
                    status_dict[vi["idx"]] = "🔄 처리 중..."
                    if parallel_pipelines:
                        out = run_pipeline_parallel(
                            video_path=vi["video_path"],
                            audio_path=vi["audio_path"],
                            api_key=api_key,
                            check_audio=check_audio,
                            check_screen=check_screen,
                            sample_rate=sample_rate,
                            diff_threshold=diff_threshold,
                            screen_batch_size=batch_size,
                            screen_max_workers=screen_workers,
                            model=selected_model,
                            domain_hint=domain_hint,
                            context_window=context_window,
                            use_sentence_merge=use_sentence_merge,
                            audio_batch_size=audio_batch_size,
                            audio_max_workers=audio_workers,
                            stt_chunk_workers=stt_chunk_workers,
                            stt_cache_dir=stt_cache_dir,
                        )
                        audio_results  = out["audio"]
                        screen_results = out["screen"]
                        errors = out.get("errors", {})
                        stt_cached = out.get("stt_cached", False)
                    else:
                        audio_results, screen_results = [], []
                        errors = {}
                        stt_cached = False
                        if check_screen:
                            try:
                                frames = extract_and_filter_frames(
                                    vi["video_path"], sample_rate, diff_threshold)
                                if frames:
                                    screen_results = spell_check_frames(
                                        frames, api_key,
                                        batch_size=batch_size,
                                        model=selected_model,
                                        max_workers=screen_workers,
                                    )
                            except Exception as e:
                                errors["screen"] = str(e)
                        if check_audio:
                            try:
                                if extract_audio(vi["video_path"], vi["audio_path"]):
                                    segments = transcribe_audio(
                                        vi["audio_path"], api_key,
                                        domain_hint=domain_hint,
                                        max_workers=stt_chunk_workers,
                                    )
                                    if segments:
                                        audio_results = spell_check_segments(
                                            segments, api_key,
                                            context_window=context_window,
                                            model=selected_model,
                                            use_sentence_merge=use_sentence_merge,
                                            batch_size=audio_batch_size,
                                            max_workers=audio_workers,
                                        )
                                        audio_results = attach_frames_to_audio_results(
                                            audio_results, vi["video_path"])
                                else:
                                    errors["audio"] = "오디오 추출 실패"
                            except Exception as e:
                                errors["audio"] = str(e)

                    err_note = ""
                    if errors:
                        err_note = " ⚠️ " + ", ".join(f"{k}: {v[:40]}" for k, v in errors.items())
                    cache_badge = " ⚡캐시" if stt_cached else ""
                    status_dict[vi["idx"]] = (
                        f"✅ 완료{cache_badge} — 🎧 {len(audio_results)}건 / "
                        f"🖼️ {len(screen_results)}건{err_note}"
                    )

                    return {
                        "idx": vi["idx"],
                        "name": vi["name"],
                        "filename": vi["filename"],
                        "audio": audio_results,
                        "screen": screen_results,
                        "errors": errors,
                        "stt_cached": stt_cached,
                        "is_fatal": False,
                    }
                except Exception as e:
                    status_dict[vi["idx"]] = f"❌ 실패: {str(e)[:100]}"
                    return {
                        "idx": vi["idx"],
                        "name": vi["name"],
                        "filename": vi["filename"],
                        "audio": [],
                        "screen": [],
                        "errors": {"fatal": str(e)},
                        "stt_cached": False,
                        "is_fatal": True,
                    }

            # ── 4) 영상 간 병렬 실행 ──────────────────────
            results_by_idx = {}
            done_count = 0
            if parallel_videos <= 1 or len(video_infos) == 1:
                for vi in video_infos:
                    results_by_idx[vi["idx"]] = _process_one(vi)
                    done_count += 1
                    overall_prog.progress(
                        done_count / len(video_infos),
                        text=f"{done_count} / {len(video_infos)}개 영상 처리 완료",
                    )
                    _render_status()
            else:
                with ThreadPoolExecutor(max_workers=parallel_videos) as ex:
                    futures = {ex.submit(_process_one, vi): vi["idx"] for vi in video_infos}
                    for fut in as_completed(futures):
                        idx = futures[fut]
                        try:
                            results_by_idx[idx] = fut.result()
                        except Exception as e:
                            results_by_idx[idx] = {
                                "idx": idx, "name": "-", "filename": "-",
                                "audio": [], "screen": [],
                                "errors": {"fatal": str(e)},
                                "stt_cached": False, "is_fatal": True,
                            }
                        done_count += 1
                        overall_prog.progress(
                            done_count / len(video_infos),
                            text=f"{done_count} / {len(video_infos)}개 영상 처리 완료",
                        )
                        _render_status()

            # 원본 순서대로 정렬
            all_video_results = [results_by_idx[vi["idx"]] for vi in video_infos]

            elapsed = _time.time() - _t0
            overall_prog.progress(1.0, text=f"✅ 전체 완료 ({elapsed:.1f}초)")

            success_count = sum(1 for r in all_video_results if not r.get("is_fatal"))
            fail_count = len(all_video_results) - success_count
            cache_hits = sum(1 for r in all_video_results if r.get("stt_cached"))

            summary_msg = (
                f"🎉 {success_count}/{len(all_video_results)}개 성공 · "
                f"소요 시간 **{elapsed:.1f}초**"
            )
            if cache_hits:
                summary_msg += f" · ⚡ STT 캐시 적중 {cache_hits}개"
            if fail_count:
                summary_msg += f" · ❌ 실패 {fail_count}개"
                st.error(summary_msg)
                # 실패 영상의 상세 에러 표시
                for r in all_video_results:
                    if r.get("is_fatal"):
                        st.caption(f"❌ **{r['filename']}**: {r['errors'].get('fatal', '알 수 없는 오류')}")
            else:
                st.success(summary_msg)

        finally:
            # ── 임시 작업 디렉토리 확실히 정리 ────────────
            import shutil as _shutil
            try:
                _shutil.rmtree(work_dir, ignore_errors=True)
            except Exception:
                pass

        # 결과가 하나도 없으면 (파일 준비 실패 등) 여기서 중단
        if not all_video_results:
            st.stop()

        # ── 전체 결과 출력 ────────────────────────────
        st.divider()
        st.subheader("📋 교정 결과")

        # 영상별 탭
        video_tab_labels = [
            f"📹 {r['name']} ({len(r['audio'])+len(r['screen'])}건)"
            for r in all_video_results
        ]
        video_tabs = st.tabs(video_tab_labels)

        for v_tab, vr in zip(video_tabs, all_video_results):
            with v_tab:
                audio_results  = vr["audio"]
                screen_results = vr["screen"]

                # 요약 지표
                c1, c2, c3 = st.columns(3)
                c1.metric("전체 교정 건수", f"{len(audio_results)+len(screen_results)}건")
                c2.metric("🎧 음성 대본",   f"{len(audio_results)}건")
                c3.metric("🖼️ 화면 텍스트", f"{len(screen_results)}건")
                st.markdown("")

                # 음성/화면 탭
                inner_labels = []
                if check_audio:  inner_labels.append(f"🎧 음성 대본 ({len(audio_results)}건)")
                if check_screen: inner_labels.append(f"🖼️ 화면 텍스트 ({len(screen_results)}건)")
                inner_tabs = st.tabs(inner_labels)
                inner_idx = 0

                if check_audio:
                    with inner_tabs[inner_idx]:
                        st.markdown("##### 음성에서 발견된 맞춤법·오타 오류")
                        render_result_cards(audio_results, badge_type="audio")
                    inner_idx += 1

                if check_screen:
                    with inner_tabs[inner_idx]:
                        st.markdown("##### 화면 텍스트에서 발견된 맞춤법·오타 오류")
                        render_result_cards(screen_results, badge_type="screen")

                # 엑셀 다운로드
                if audio_results or screen_results:
                    st.divider()
                    excel_bytes = make_excel(
                        audio_results, screen_results,
                        reviewer, lesson_name, lesson_num, review_date
                    )
                    base = vr["name"]
                    fname = f"{lesson_num}차시_{base}_교정결과.xlsx" if lesson_name else f"{base}_교정결과.xlsx"
                    st.download_button(
                        label=f"📥 {vr['filename']} 교정 결과 다운로드",
                        data=excel_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                        key=f"dl_{vr['name']}"
                    )
                    st.info("💡 엑셀 파일에는 **음성 대본 / 화면 텍스트 / 통합** 시트가 포함됩니다.")

        # ── 전체 통합 다운로드 (영상 2개 이상일 때만 표시) ──
        if len(all_video_results) >= 2:
            total_cnt = sum(len(vr["audio"]) + len(vr["screen"]) for vr in all_video_results)
            if total_cnt > 0:
                st.divider()
                st.markdown("### 📦 전체 통합 다운로드")
                st.caption(f"총 {len(all_video_results)}개 영상 · {total_cnt}건 교정 항목을 하나의 엑셀 파일로 다운로드합니다.")
                combined_bytes = make_excel_combined(
                    all_video_results, reviewer, lesson_name, lesson_num, review_date
                )
                fname_combined = f"{lesson_name}_전체통합_교정결과.xlsx" if lesson_name else "전체통합_교정결과.xlsx"
                st.download_button(
                    label="📥 전체 통합 엑셀 다운로드 (.xlsx)",
                    data=combined_bytes,
                    file_name=fname_combined,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key="dl_combined"
                )
                st.info("💡 통합 파일 구조: [영상명_음성] [영상명_화면] 시트 × 영상 수 + [📋 전체통합] 시트")
