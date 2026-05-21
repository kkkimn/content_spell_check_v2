from pathlib import Path

path = Path("app_video.py")
raw = path.read_bytes()

for enc in ("utf-8", "cp949", "euc-kr"):
    try:
        text = raw.decode(enc)
        break
    except UnicodeDecodeError:
        text = None

if text is None:
    text = raw.decode("utf-8", errors="replace")

path.write_text(text, encoding="utf-8", newline="
")
print("app_video.py encoding normalized to UTF-8")
