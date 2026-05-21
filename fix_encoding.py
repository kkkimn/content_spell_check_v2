with open('app_video.py', 'rb') as f:
    raw_lines = f.readlines()

out_lines = []
for i, line_b in enumerate(raw_lines):
    # lines 267-275 are indices 266-274
    if 266 <= i <= 274:
        continue
    # lines 444-454 are indices 443-453
    if 443 <= i <= 453:
        continue
    
    # decode robustly
    try:
        text = line_b.decode('utf-8')
    except UnicodeDecodeError:
        try:
            text = line_b.decode('cp949')
        except UnicodeDecodeError:
            text = line_b.decode('utf-8', errors='replace')
            
    out_lines.append(text)

with open('app_video.py', 'w', encoding='utf-8') as f:
    f.writelines(out_lines)
