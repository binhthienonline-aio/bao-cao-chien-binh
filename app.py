import io, os, re, zipfile, tempfile, unicodedata, math
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import pandas as pd
import numpy as np
import cv2
import pytesseract
import gdown
from PIL import Image, ImageOps, ImageEnhance, ImageDraw, ImageFont
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Báo cáo Chiến Binh V8", page_icon="📊", layout="wide")
st.markdown("""
<style>
.block-container{max-width:1600px;padding-top:1.1rem}
[data-testid="stSidebar"]{background:#f3f6fa}
.hero{background:linear-gradient(110deg,#063b70,#0879c5);padding:22px 26px;border-radius:14px;color:white;margin-bottom:16px}
.hero h1{color:#fff;margin:0;font-size:26px}.hero p{color:#e5f2ff;margin:6px 0 0}
div.stButton>button,div.stDownloadButton>button{border-radius:8px;font-weight:650;min-height:40px}
</style>
""", unsafe_allow_html=True)
st.markdown(
    '<div class="hero"><h1>📊 BÁO CÁO SỨC KHỎE CHIẾN BINH — V8</h1>'
    '<p>Chỉ đọc ảnh tháng 11 • Lũy kế đọc riêng THỰC HIỆN + TARGET • Phân bổ theo tỷ trọng Dashboard</p></div>',
    unsafe_allow_html=True
)

EMPLOYEES = [
    ("137501", "Võ Văn Lượng"),
    ("137502", "Nguyễn Ngọc Tân"),
    ("137506", "Lưu Thị Tuyết Nhi"),
    ("170206", "Lê Thị Thu Hiền"),
    ("181846", "Lưu Văn Nhật"),
    ("21512", "Hà Thị Minh"),
    ("269200", "Võ Quốc Cường"),
    ("281296", "Phạm Xuân Toàn"),
    ("82654", "Nguyễn Văn Quyền"),
]

def norm(text):
    text = unicodedata.normalize("NFD", str(text or "").lower())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()

def compact(text):
    return norm(text).replace(" ", "")

def base_name(path):
    return Path(str(path).replace("\\", "/")).name

def image_bytes_to_bgr(raw):
    arr = cv2.imdecode(np.frombuffer(raw, np.uint8), cv2.IMREAD_COLOR)
    if arr is None:
        raise ValueError("Không mở được ảnh.")
    return arr

def parse_numeric(text):
    if text is None:
        return None
    s = str(text).strip().replace(" ", "")
    # Source image uses comma as thousands separator and dot as decimal.
    s = s.replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None

def unzip_images(raw):
    images = []
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        bad = zf.testzip()
        if bad:
            raise ValueError(f"File ZIP bị lỗi tại {bad}.")
        for info in zf.infolist():
            if info.is_dir():
                continue
            if Path(info.filename).suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
                images.append((base_name(info.filename), zf.read(info.filename)))
    return images

def load_public_drive(url):
    tmp = tempfile.mkdtemp(prefix="drive_v5_")
    if "/folders/" in url:
        downloaded = gdown.download_folder(url=url, output=tmp, quiet=True, use_cookies=False)
        if not downloaded:
            raise ValueError("Không tải được thư mục Drive. Hãy bật quyền 'Bất kỳ ai có đường liên kết'.")
        out = []
        for path in Path(tmp).rglob("*"):
            if path.is_file() and path.suffix.lower() == ".zip":
                out.extend(unzip_images(path.read_bytes()))
            elif path.is_file() and path.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
                out.append((path.name, path.read_bytes()))
        return out
    target = os.path.join(tmp, "source.zip")
    downloaded = gdown.download(url=url, output=target, quiet=True)
    if not downloaded:
        raise ValueError("Không tải được link Drive. Kiểm tra quyền xem công khai.")
    if Path(downloaded).suffix.lower() == ".zip":
        return unzip_images(Path(downloaded).read_bytes())
    raise ValueError("Link Drive cần trỏ tới file ZIP ảnh hoặc thư mục ảnh công khai.")

def identify_sources(images):
    target = []
    employee_files = {}
    others = []
    for name, raw in images:
        n = norm(name)
        if "luy" in n and "ke" in n:
            target.append((name, raw))
            continue
        code = next((c for c, _ in EMPLOYEES if c in name), None)
        if code:
            employee_files.setdefault(code, []).append((name, raw))
        else:
            others.append(name)
    # Nếu Drive còn lưu cả ảnh Lũy kế cũ và mới, ưu tiên ảnh chụp màn hình điện thoại
    # có tỷ lệ dọc lớn; không lấy ảnh bảng Excel cũ làm nguồn THỰC HIỆN.
    def target_rank(item):
        try:
            img = Image.open(io.BytesIO(item[1]))
            w, h = img.size
            return (h / max(w, 1), h * w)
        except Exception:
            return (0, 0)
    target.sort(key=target_rank, reverse=True)
    missing = [c for c, _ in EMPLOYEES if c not in employee_files]
    duplicates = [c for c, v in employee_files.items() if len(v) > 1]
    return target, employee_files, missing, duplicates, others

def _runs(binary_array):
    runs = []
    start = None
    for i, value in enumerate(binary_array):
        if value and start is None:
            start = i
        elif not value and start is not None:
            runs.append((start, i - 1, i - start))
            start = None
    if start is not None:
        runs.append((start, len(binary_array) - 1, len(binary_array) - start))
    return runs

def find_target_table_geometry(bgr):
    h, w = bgr.shape[:2]
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    x0, x1 = int(w * 0.015), int(w * 0.985)
    scores = (gray[:, x0:x1] < 180).sum(axis=1)
    active = scores > 0.85 * (x1 - x0)
    runs = _runs(active)
    thin_points = [a for a, b, length in runs if length <= 3 and a > int(h * 0.06)]
    # Find the longest run of evenly spaced table row separators.
    best_seq = []
    for i, point in enumerate(thin_points):
        seq = [point]
        for nxt in thin_points[i + 1:]:
            gap = nxt - seq[-1]
            if 0.017 * h <= gap <= 0.030 * h:
                seq.append(nxt)
            elif gap > 0.030 * h:
                break
        if len(seq) > len(best_seq):
            best_seq = seq
    if len(best_seq) < 3:
        raise ValueError("Không tìm thấy lưới dòng của ảnh Lũy kế.")
    # Vertical table separators identify the group-name and TARGET columns.
    y0, y1 = best_seq[0], best_seq[-1]
    vscore = (gray[y0:y1 + 1, :] < 180).sum(axis=0)
    vactive = vscore > 0.82 * (y1 - y0 + 1)
    vlines = []
    for a, b, length in _runs(vactive):
        if length <= 4 and a > int(w * 0.01) and b < int(w * 0.995):
            vlines.append(int((a + b) / 2))
    if len(vlines) >= 8:
        # Observed sheet structure: STT, Group, Expected %, Target/NV, Luy ke, TARGET, Con thieu, Tang/Giam.
        group_x = (vlines[1], vlines[2])
        target_x = (vlines[5], vlines[6])
    else:
        group_x = (int(w * 0.075), int(w * 0.354))
        target_x = (int(w * 0.688), int(w * 0.778))
    return best_seq, group_x, target_x


def _ocr_words_in_crop(pil, box, lang="eng", config="--psm 6", scale=3):
    """OCR từ một cột riêng; trả về token cùng tọa độ Y trên ảnh gốc."""
    x0, y0, x1, y1 = box
    crop = pil.crop((x0, y0, x1, y1))
    crop = crop.resize((max(1, crop.width * scale), max(1, crop.height * scale)),
                       Image.Resampling.LANCZOS)
    data = pytesseract.image_to_data(
        crop, lang=lang, config=config, output_type=pytesseract.Output.DICT
    )
    out = []
    for i, raw_word in enumerate(data["text"]):
        word = str(raw_word).strip()
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1
        if not word or conf < 10:
            continue
        top = y0 + float(data["top"][i]) / scale
        height = float(data["height"][i]) / scale
        left = x0 + float(data["left"][i]) / scale
        width = float(data["width"][i]) / scale
        out.append({
            "text": word, "cy": top + height / 2,
            "cx": left + width / 2, "conf": conf
        })
    return out


def parse_target_image(raw):
    """
    Đọc ảnh Lũy kế dạng bảng điện thoại mới:
    - Cột chương trình
    - THỰC HIỆN
    - TARGET
    - Các cột % phía bên phải chỉ dùng để nhận diện bố cục, không dùng tính doanh số.
    Mỗi dòng được neo theo tọa độ OCR của cột THỰC HIỆN để tránh lệch cột/lệch dòng.
    """
    bgr = image_bytes_to_bgr(raw)
    pil = Image.fromarray(cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB))
    h, w = bgr.shape[:2]

    # Ảnh bảng điện thoại: cột cố định theo tỷ lệ chiều rộng, không dùng lưới của Excel tháng 10.
    if h / max(w, 1) >= 2.0 and w >= 400:
        y_top, y_bottom = int(h * 0.185), int(h * 0.992)
        name_right = int(w * 0.372)
        actual_right = int(w * 0.536)
        target_right = int(w * 0.698)

        actual_words = _ocr_words_in_crop(
            pil, (int(w * 0.372), y_top, actual_right, y_bottom),
            lang="eng", config="--psm 6 -c tessedit_char_whitelist=0123456789,.-"
        )
        target_words = _ocr_words_in_crop(
            pil, (actual_right, y_top, target_right, y_bottom),
            lang="eng", config="--psm 6 -c tessedit_char_whitelist=0123456789,.-"
        )
        name_words = _ocr_words_in_crop(
            pil, (0, y_top, name_right, y_bottom),
            lang="vie+eng", config="--psm 6"
        )

        # Một token số THỰC HIỆN là mốc của một dòng dữ liệu.
        anchors = []
        for token in actual_words:
            value = parse_numeric(token["text"])
            if value is not None:
                anchors.append((token["cy"], value, token["text"], token["conf"]))
        anchors.sort(key=lambda x: x[0])

        # Gộp token cùng một dòng nếu OCR tách một số thành nhiều mảnh.
        merged = []
        for item in anchors:
            if merged and abs(item[0] - merged[-1][0]) < 10:
                # Giữ token có độ tin cậy cao hơn, tránh tự ghép hai con số.
                if item[3] > merged[-1][3]:
                    merged[-1] = item
            else:
                merged.append(item)
        anchors = merged

        if len(anchors) != 39:
            raise ValueError(
                f"Ảnh Lũy kế phải có 39 chương trình nhưng OCR chỉ nhận diện {len(anchors)} giá trị THỰC HIỆN. "
                "V8 dừng tại đây, không tự bù hoặc bỏ dòng. Hãy kiểm tra ảnh Drive đã thay đúng ảnh mới chưa."
            )

        records = []
        for idx, (cy, actual, actual_raw, actual_conf) in enumerate(anchors):
            prev_mid = ((anchors[idx - 1][0] + cy) / 2) if idx else cy - 24
            next_mid = ((cy + anchors[idx + 1][0]) / 2) if idx + 1 < len(anchors) else cy + 24

            # OCR tên riêng trong đúng vùng dòng; tốt hơn ghép token toàn ảnh khi tên xuống 2 dòng.
            row_crop = pil.crop((
                int(w * 0.065), max(0, int(prev_mid)),
                name_right, min(h, int(next_mid))
            ))
            row_crop = row_crop.resize((
                max(1, row_crop.width * 3), max(1, row_crop.height * 3)
            ), Image.Resampling.LANCZOS)
            label_text = pytesseract.image_to_string(
                row_crop, lang="vie+eng", config="--psm 6"
            )
            label = re.sub(r"\s+", " ", label_text).strip()
            label = re.sub(r"^(?:[+*#©®|]+\s*)+", "", label)
            label = re.sub(r"\b(?:DT|SL|ICT|DV|or|opr|st|si|srL)\b", "", label, flags=re.I)
            label = re.sub(r"\s+", " ", label).strip(" -|+")

            # TARGET chỉ lấy token nằm đúng vùng cột TARGET và đúng dải Y của dòng.
            row_target = [
                t for t in target_words
                if prev_mid <= t["cy"] < next_mid
                and parse_numeric(t["text"]) is not None
            ]
            row_target.sort(key=lambda t: (abs(t["cy"] - cy), -t["conf"]))
            target = parse_numeric(row_target[0]["text"]) if row_target else None
            target_raw = row_target[0]["text"] if row_target else ""

            # Nhóm Đồng hồ & Smartwatch trong ảnh không có TARGET.
            is_smartwatch = "smartwatch" in norm(label)
            if is_smartwatch:
                target = None
                target_raw = "Không có TARGET trên ảnh"

            status_parts = []
            if not label:
                status_parts.append("Cần kiểm tra tên nhóm")
            if not is_smartwatch and target is None:
                status_parts.append("Cần kiểm tra TARGET")
            if actual_conf < 35:
                status_parts.append("OCR THỰC HIỆN độ tin cậy thấp")
            if not status_parts:
                status_parts.append("Đã đọc — cần đối chiếu")
            records.append({
                "STT": idx + 1,
                "Nhóm hàng": label,
                "THỰC HIỆN": actual,
                "TARGET siêu thị": target,
                "OCR tên": label,
                "OCR số": f"THỰC HIỆN={actual_raw}; TARGET={target_raw}",
                "Trạng thái": "; ".join(status_parts)
            })

        return pd.DataFrame(records)

    # Không dùng ảnh bảng Excel cũ làm nguồn tháng 11 vì ảnh đó không có cột THỰC HIỆN.
    # Nếu ảnh không phải màn hình điện thoại mới, báo rõ để người dùng thay đúng ảnh.
    raise ValueError(
        f"Ảnh Lũy kế '{w}×{h}' không đúng định dạng mới. Hãy thay bằng ảnh chụp màn hình điện thoại "
        "có đủ cột CHƯƠNG TRÌNH, THỰC HIỆN, TARGET, % HT. Không dùng ảnh bảng Excel tháng 10."
    )
    row_lines, group_x, target_x = find_target_table_geometry(bgr)
    records = []
    for i in range(len(row_lines) - 1):
        y0, y1 = row_lines[i], row_lines[i + 1]
        name_crop = pil.crop((group_x[0] + 3, y0 + 3, group_x[1] - 3, y1 - 3))
        name_crop = name_crop.resize((name_crop.width * 2, name_crop.height * 2))
        name_text = pytesseract.image_to_string(name_crop, lang="vie+eng", config="--psm 6")
        name = re.sub(r"\s+", " ", name_text).strip(" \n\t-:")
        target_crop = pil.crop((target_x[0] + 3, y0 + 3, target_x[1] - 3, y1 - 3))
        target_text = pytesseract.image_to_string(
            target_crop, lang="eng",
            config="--psm 7 -c tessedit_char_whitelist=0123456789,."
        )
        target = parse_numeric(target_text)
        records.append({
            "STT": i + 1, "Nhóm hàng": name, "THỰC HIỆN": None,
            "TARGET siêu thị": target, "OCR tên": name,
            "OCR số": f"TARGET={target_text.strip()}",
            "Trạng thái": "Cần kiểm tra THỰC HIỆN — ảnh không đúng định dạng mới"
        })
    return pd.DataFrame(records)

def card_components(bgr):
    h, w = bgr.shape[:2]
    y0, y1 = int(h * 0.50), int(h * 0.74)
    x0, x1 = int(w * 0.015), int(w * 0.985)
    roi = bgr[y0:y1, x0:x1]
    # Dashboard cards are white on the light-pink THI ĐUA panel.
    mask = cv2.inRange(roi, (250, 250, 250), (255, 255, 255))
    mask = cv2.morphologyEx(
        mask, cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    )
    count, labels, stats, centers = cv2.connectedComponentsWithStats(mask)
    boxes = []
    for x, y, bw, bh, area in stats[1:]:
        if (area > 0.0015 * w * h
            and 0.14 * w <= bw <= 0.22 * w
            and 0.016 * h <= bh <= 0.050 * h):
            boxes.append((int(x + x0), int(y + y0), int(bw), int(bh)))
    boxes.sort(key=lambda b: (b[1], b[0]))
    return boxes

def card_ocr(bgr, box):
    x, y, w, h = box
    H, W = bgr.shape[:2]
    # Whole card OCR is used for the category label only.
    crop = bgr[max(0, y - 5):min(H, y + h + 5), max(0, x - 5):min(W, x + w + 5)]
    crop_up = cv2.resize(crop, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    pil_full = Image.fromarray(cv2.cvtColor(crop_up, cv2.COLOR_BGR2RGB))
    try:
        label_text = pytesseract.image_to_string(pil_full, lang="eng", config="--psm 6", timeout=8)
    except RuntimeError:
        label_text = ""

    # Isolate the colored percentage glyphs in the top-right corner.
    # This removes pale grey diagonal watermark strokes that OCR can mistake for +/- or extra digits.
    ch, cw = crop.shape[:2]
    num_crop = crop[:max(12, int(ch * 0.54)), int(cw * 0.60):]
    hsv = cv2.cvtColor(num_crop, cv2.COLOR_BGR2HSV)
    colored = ((hsv[:, :, 1] > 70) & (hsv[:, :, 2] > 60)).astype(np.uint8) * 255
    binary = Image.fromarray(255 - colored)
    numeric_texts = []
    parsed = []
    for psm in (7, 10, 6):
        try:
            txt = pytesseract.image_to_string(
                binary, lang="eng",
                config=f"--psm {psm} -c tessedit_char_whitelist=0123456789%.-",
                timeout=5
            ).strip()
        except RuntimeError:
            txt = ""
        numeric_texts.append(txt)
        match = re.search(r"(-?\s*\d{1,3})\s*%", txt)
        if match:
            try:
                parsed.append(float(re.sub(r"\s+", "", match.group(1))))
            except ValueError:
                pass
    pct = None
    if parsed:
        # Prefer consensus across OCR modes; otherwise use the most frequent reading.
        counts = {}
        for value in parsed:
            counts[value] = counts.get(value, 0) + 1
        pct = sorted(counts.items(), key=lambda x: (x[1], -abs(x[0])))[-1][0]
        if counts[pct] == 1 and len(parsed) > 1:
            pct = None  # disagreement -> require manual review

    label = re.sub(r"-?\s*\d{1,3}\s*%", " ", label_text)
    label = re.sub(r"[\n\r]+", " ", label)
    label = re.sub(r"[^A-Za-z0-9À-ỹ/&+ -]", " ", label)
    label = re.sub(r"\s+", " ", label).strip(" -_")
    raw_text = label_text.strip() + " | NUMBER: " + " / ".join(x for x in numeric_texts if x)
    return label, pct, raw_text

def match_group(card_label, groups):
    c_norm, c_compact = norm(card_label), compact(card_label)
    if not c_norm:
        return None, 0.0, "Không đọc được tên thẻ"
    c_tokens = set(c_norm.split())
    scored = []
    for idx, group in enumerate(groups):
        g_norm, g_compact = norm(group), compact(group)
        g_tokens = set(g_norm.split())
        if not g_norm:
            continue
        if c_norm == g_norm or c_compact == g_compact:
            score = 1000 + len(g_tokens)
        elif g_tokens and g_tokens.issubset(c_tokens):
            # Prefer a more specific group when labels overlap, e.g. Sim vs Sim DMX.
            score = 900 + len(g_tokens)
        elif c_tokens and c_tokens.issubset(g_tokens) and len(c_tokens) >= 2:
            score = 800 + len(c_tokens)
        else:
            score = max(
                fuzz.token_set_ratio(c_norm, g_norm),
                fuzz.ratio(c_compact, g_compact),
                fuzz.partial_ratio(c_compact, g_compact) * 0.90,
            )
        scored.append((score, len(g_tokens), idx, group))
    scored.sort(reverse=True)
    if not scored or scored[0][0] < 76:
        return None, (scored[0][0] if scored else 0), "Tên nhóm chưa ghép được"
    best = scored[0]
    second = scored[1] if len(scored) > 1 else (0, 0, -1, "")
    if best[0] < 100 and best[0] - second[0] < 4:
        return None, best[0], f"Nhóm chưa chắc chắn: {best[3]} / {second[3]}"
    return best[2], best[0], "Đã ghép"

def extract_no_data(raw, groups):
    bgr = image_bytes_to_bgr(raw)
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    im = Image.fromarray(rgb)
    h, w = bgr.shape[:2]
    crop = im.crop((int(w * 0.02), int(h * 0.68), int(w * 0.985), int(h * 0.76)))
    crop = ImageOps.grayscale(crop.resize((crop.width * 2, crop.height * 2)))
    crop = ImageEnhance.Contrast(crop).enhance(1.5)
    text = pytesseract.image_to_string(crop, lang="vie+eng", config="--psm 6")
    lines = text.splitlines()
    start_line = None
    start_idx = None
    for i, line in enumerate(lines):
        if "chua co so lieu" in norm(line):
            start_line = line
            start_idx = i
            break
    if start_idx is None:
        return None, {}, text

    head = norm(start_line)
    count_match = re.search(r"chua co so lieu\s*\(?\s*(\d+)\s*\)?", head)
    expected = int(count_match.group(1)) if count_match else None
    colon = start_line.find(":")
    tail_lines = [start_line[colon + 1:] if colon >= 0 else ""]
    for line in lines[start_idx + 1:]:
        if "sieu thi khong co target" in norm(line):
            break
        tail_lines.append(line)
    raw_tail = " ".join(tail_lines)
    # Preserve separators before accent normalization. Some composite categories are split
    # into sub-parts (e.g. Máy lọc KK - Hút bụi - Hút ẩm); duplicate category matches collapse.
    segments = [seg.strip(" :;,.\t") for seg in re.split(r"\s*[-–—]\s*|\s*:\s*", raw_tail) if seg.strip(" :;,.\t")]
    matched = {}
    evidence = {}
    for segment in segments:
        idx, score, status = match_group(segment, groups)
        if idx is not None and idx not in matched:
            matched[idx] = 0.0
            evidence[idx] = segment
    return expected, matched, text

def parse_dashboard(raw, groups, code, employee_name):
    bgr = image_bytes_to_bgr(raw)
    boxes = card_components(bgr)
    cards = []
    values = {}
    duplicate_groups = set()
    seq_values = []
    for pos, box in enumerate(boxes, 1):
        label, pct, raw_text = card_ocr(bgr, box)
        idx, score, match_status = match_group(label, groups)
        status = match_status
        if pct is None:
            status = "Không đọc được tỷ lệ"
        if idx is not None and pct is not None:
            if idx in values:
                duplicate_groups.add(idx)
                status = "Trùng nhóm — cần kiểm tra"
            else:
                values[idx] = pct
        else:
            idx = None
        if pct is not None:
            seq_values.append((pos, pct, label))
        cards.append({
            "Mã NV": code,
            "Nhân viên": f"{employee_name} - {code}",
            "STT thẻ": pos,
            "Tên OCR": label,
            "Nhóm ghép": (groups[idx] if idx is not None else ""),
            "Tỷ lệ (%)": pct,
            "Điểm ghép": round(float(score), 1),
            "Trạng thái": status,
            "OCR thô": raw_text
        })
    expected_zero, zero_map, zero_raw = extract_no_data(raw, groups)
    # Zero from the dashboard's explicit "Chưa có số liệu" list only.
    for idx in zero_map:
        if idx not in values and idx not in duplicate_groups:
            values[idx] = 0.0
    conflicts = sorted(set(values).intersection(duplicate_groups))
    for idx in conflicts:
        values.pop(idx, None)
    order_issues = []
    previous = None
    for pos, pct, label in seq_values:
        if previous is not None and pct > previous[1] + 1:
            order_issues.append({"STT thẻ": pos, "Tỷ lệ": pct, "Thẻ trước": previous[1], "Tên OCR": label})
        previous = (pos, pct)
    return {
        "employee": f"{employee_name} - {code}",
        "code": code,
        "card_count": len(boxes),
        "values": values,
        "zero_expected": expected_zero,
        "zero_matched": len(zero_map),
        "zero_groups": list(zero_map.keys()),
        "cards": cards,
        "duplicates": sorted(duplicate_groups),
        "order_issues": order_issues,
        "zero_raw": zero_raw
    }

def build_excel(groups, sales, rates, report_date, staff, audit_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Báo cáo"
    n = len(groups)
    last = n + 1
    ws.cell(1, 1, n)
    for col, group in enumerate(groups, 2):
        ws.cell(1, col, group)
    ws.cell(2, 1, "TỔNG")
    for j, group in enumerate(groups, 2):
        col_vals = pd.to_numeric(sales[group], errors="coerce")
        ws.cell(2, j, float(col_vals.sum()) if col_vals.notna().all() else None)
    for row_idx in range(len(sales)):
        ws.cell(row_idx + 3, 1, sales.iloc[row_idx]["Nhân viên"])
        for col_idx, group in enumerate(groups, 2):
            v = sales.iloc[row_idx][group]
            ws.cell(row_idx + 3, col_idx, float(v) if pd.notna(v) else None)

    yellow, orange, blue = "FFF200", "FFC000", "DDEBF7"
    thin = Side(style="thin", color="202020")
    for col in range(1, last + 1):
        c = ws.cell(1, col)
        c.fill = PatternFill("solid", fgColor=yellow)
        c.font = Font(name="Arial", size=8, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        c = ws.cell(2, col)
        c.fill = PatternFill("solid", fgColor=orange)
        c.font = Font(name="Arial", size=9, bold=True, color="000000")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if col > 1:
            c.number_format = "#,##0.00"
    for row in range(3, 12):
        for col in range(1, last + 1):
            c = ws.cell(row, col)
            c.fill = PatternFill("solid", fgColor=blue)
            c.font = Font(name="Arial", size=9, color="000000")
            c.alignment = Alignment(horizontal="left" if col == 1 else "center", vertical="center")
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if col > 1:
                c.number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 29
    for col in range(2, last + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.row_dimensions[1].height = 64
    ws.row_dimensions[2].height = 22
    for row in range(3, 12):
        ws.row_dimensions[row].height = 19
    ws.freeze_panes = "B3"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{get_column_letter(last)}11"

    audit = wb.create_sheet("Kiểm tra OCR")
    audit.append(["Ngày báo cáo", report_date])
    audit.append(["Số nhân viên", staff])
    audit.append(["Số nhóm hàng", n])
    audit.append(["Ô tỷ lệ chưa xác minh", int(rates.isna().sum().sum())])
    audit.append([])
    audit.append(["Mã NV", "Nhân viên", "STT thẻ", "Tên OCR", "Nhóm ghép", "Tỷ lệ", "Điểm ghép", "Trạng thái", "OCR thô"])
    for row in audit_df.itertuples(index=False):
        audit.append(list(row))
    for c in audit[7]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for col, width in {"A": 14, "B": 28, "C": 12, "D": 32, "E": 38, "F": 12, "G": 12, "H": 28, "I": 70}.items():
        audit.column_dimensions[col].width = width
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def font_file(size, bold=False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"
    ]
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()

def wrap_text_px(draw, text, font, max_width):
    words = str(text).split()
    lines, current = [], ""
    for word in words:
        candidate = word if not current else current + " " + word
        if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width or not current:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]

def build_png(groups, matrix, report_date):
    left_w, col_w = 180, 54
    header_h, total_h, row_h = 88, 28, 30
    width = left_w + col_w * len(groups) + 1
    height = header_h + total_h + row_h * len(matrix) + 1
    img = Image.new("RGB", (width, height), "white")
    d = ImageDraw.Draw(img)
    f_header = font_file(10, True)
    f_total = font_file(11, True)
    f_row = font_file(10, False)
    f_num = font_file(9, False)
    grid = (20, 20, 20)
    yellow, orange, blue = (255, 242, 0), (255, 192, 0), (221, 235, 247)

    # Header band
    d.rectangle((0, 0, left_w, header_h), fill=yellow, outline=grid, width=1)
    d.text((left_w / 2, header_h / 2), str(len(groups)), font=f_header, fill="black", anchor="mm")
    for j, group in enumerate(groups):
        x0 = left_w + j * col_w
        d.rectangle((x0, 0, x0 + col_w, header_h), fill=yellow, outline=grid, width=1)
        lines = wrap_text_px(d, group, f_header, col_w - 6)
        line_h = d.textbbox((0, 0), "Ag", font=f_header)[3] + 1
        y = max(2, (header_h - len(lines) * line_h) / 2)
        for line in lines:
            d.text((x0 + col_w / 2, y), line, font=f_header, fill="black", anchor="mt")
            y += line_h

    # Total row
    y0 = header_h
    d.rectangle((0, y0, left_w, y0 + total_h), fill=orange, outline=grid, width=1)
    d.text((8, y0 + total_h / 2), "TỔNG", font=f_total, fill="black", anchor="lm")
    for j, group in enumerate(groups):
        x0 = left_w + j * col_w
        d.rectangle((x0, y0, x0 + col_w, y0 + total_h), fill=orange, outline=grid, width=1)
        vals = pd.to_numeric(matrix[group], errors="coerce")
        total = vals.sum() if vals.notna().all() else None
        txt = f"{total:,.2f}" if total is not None else ""
        d.text((x0 + col_w / 2, y0 + total_h / 2), txt, font=f_num, fill="black", anchor="mm")

    # Employee rows
    for i, row in matrix.iterrows():
        y = header_h + total_h + i * row_h
        d.rectangle((0, y, left_w, y + row_h), fill=blue, outline=grid, width=1)
        d.text((5, y + row_h / 2), str(row["Nhân viên"]), font=f_row, fill="black", anchor="lm")
        for j, group in enumerate(groups):
            x0 = left_w + j * col_w
            d.rectangle((x0, y, x0 + col_w, y + row_h), fill=blue, outline=grid, width=1)
            v = row[group]
            txt = f"{float(v):,.2f}" if pd.notna(v) else ""
            d.text((x0 + col_w / 2, y + row_h / 2), txt, font=f_num, fill="black", anchor="mm")

    bio = io.BytesIO()
    img.save(bio, format="PNG", optimize=True)
    bio.seek(0)
    return bio

# ---------------- Sidebar / State ----------------
with st.sidebar:
    st.markdown("### CẤU HÌNH BÁO CÁO")
    report_date = st.text_input("Ngày báo cáo", "09/11/2026")
    staff_count = st.number_input("Số nhân viên", min_value=1, max_value=100, value=9, step=1)
    st.caption("Doanh số từng NV = THỰC HIỆN cửa hàng × tỷ lệ Dashboard của NV / tổng tỷ lệ 9 NV.")
    st.caption("Tổng doanh số 9 NV theo từng nhóm phải khớp THỰC HIỆN cửa hàng. Không dùng hệ số 110%.")

for key, default in [
    ("image_items", []), ("parsed_target", None), ("dashboard_results", None),
    ("target_editor", None), ("matrix_editor", None), ("card_audit", None),
    ("report_matrix", None)
]:
    if key not in st.session_state:
        st.session_state[key] = default

tabs = st.tabs(["① NẠP ẢNH","② ĐỌC 10 ẢNH","③ RÀ SOÁT 342 Ô","④ XEM TRƯỚC & XUẤT"])

with tabs[0]:
    st.markdown("### Chỉ sử dụng ảnh tháng 11")
    left, right = st.columns(2)
    with left:
        zip_file = st.file_uploader("Tải ZIP gồm 9 dashboard + ảnh Lũy kế", type=["zip"])
        many_files = st.file_uploader("Hoặc tải 10 ảnh trực tiếp", type=["png","jpg","jpeg","webp"], accept_multiple_files=True)
        if st.button("NẠP ẢNH", type="primary"):
            try:
                items = unzip_images(zip_file.getvalue()) if zip_file else []
                items += [(f.name, f.getvalue()) for f in (many_files or [])]
                if not items:
                    st.warning("Anh chưa chọn ZIP hoặc ảnh.")
                else:
                    st.session_state.image_items = items
                    st.session_state.parsed_target = None
                    st.session_state.dashboard_results = None
                    st.session_state.target_editor = None
                    st.session_state.matrix_editor = None
                    st.session_state.report_matrix = None
                    st.success(f"Đã nạp {len(items)} ảnh.")
            except Exception as e:
                st.error(f"Không đọc được file: {e}")
    with right:
        drive_url = st.text_input("Link Google Drive công khai (ZIP hoặc thư mục ảnh)", placeholder="Dán link chia sẻ công khai")
        if st.button("TẢI ẢNH TỪ DRIVE", disabled=not drive_url):
            try:
                items = load_public_drive(drive_url)
                if not items:
                    raise ValueError("Không tìm thấy ảnh trong Drive.")
                st.session_state.image_items = items
                st.session_state.parsed_target = None
                st.session_state.dashboard_results = None
                st.session_state.target_editor = None
                st.session_state.matrix_editor = None
                st.session_state.report_matrix = None
                st.success(f"Đã tải {len(items)} ảnh từ Drive.")
            except Exception as e:
                st.error(f"Lỗi Google Drive: {e}")

    if st.session_state.image_items:
        target_images, employee_files, missing_codes, duplicate_codes, other_files = identify_sources(st.session_state.image_items)
        a, b, c = st.columns(3)
        a.metric("Tổng ảnh", len(st.session_state.image_items))
        b.metric("Dashboard nhận diện", len(employee_files))
        c.metric("Ảnh Lũy kế", len(target_images))
        if missing_codes:
            st.warning("Thiếu dashboard mã: " + ", ".join(missing_codes))
        if duplicate_codes:
            st.warning("Có nhiều ảnh cùng mã: " + ", ".join(duplicate_codes))
        if not target_images:
            st.error("Không tìm thấy ảnh Lũy kế. Tên file cần có 'Lũy kế' hoặc 'Luy ke'.")
        if target_images and not missing_codes and not duplicate_codes:
            st.success("Đã nhận diện đủ ảnh Lũy kế và 9 mã nhân viên.")
        with st.expander("Danh sách ảnh đã nạp"):
            for name, _ in st.session_state.image_items:
                st.write("• " + name)

with tabs[1]:
    st.markdown("### Phân tích ảnh Lũy kế và 9 Dashboard")
    st.info("Tên nhóm, THỰC HIỆN và TARGET lấy từ ảnh Lũy kế tháng 11. Tỷ lệ lấy từ Dashboard. Doanh số 9 NV được phân bổ theo tỷ trọng tỷ lệ và tổng từng nhóm khớp THỰC HIỆN cửa hàng.")
    if st.button("▶ PHÂN TÍCH 10 ẢNH", type="primary", disabled=not st.session_state.image_items):
        target_images, employee_files, missing_codes, duplicate_codes, other_files = identify_sources(st.session_state.image_items)
        if not target_images:
            st.error("ZIP chưa có ảnh Lũy kế.")
        elif missing_codes:
            st.error("Thiếu dashboard mã: " + ", ".join(missing_codes))
        elif duplicate_codes:
            st.error("Trùng ảnh dashboard mã: " + ", ".join(duplicate_codes))
        else:
            try:
                with st.spinner("Đang đọc danh mục và TARGET từ ảnh Lũy kế..."):
                    target_df = parse_target_image(target_images[0][1])
                st.session_state.parsed_target = target_df
                st.session_state.target_editor = target_df.copy()
                groups = target_df["Nhóm hàng"].astype(str).tolist()
                results = {}
                all_audit = []
                progress = st.progress(0)
                status = st.empty()
                for i, (code, employee_name) in enumerate(EMPLOYEES):
                    filename, raw = employee_files[code][0]
                    status.write(f"Đang đọc {employee_name} — {code}...")
                    result = parse_dashboard(raw, groups, code, employee_name)
                    result["filename"] = filename
                    results[code] = result
                    all_audit.extend(result["cards"])
                    progress.progress((i + 1) / len(EMPLOYEES))
                st.session_state.dashboard_results = results
                st.session_state.card_audit = pd.DataFrame(all_audit)
                st.session_state.matrix_editor = None
                st.session_state.report_matrix = None
                status.success("Đã hoàn tất phân tích ảnh.")
            except Exception as e:
                st.error(f"Lỗi khi phân tích ảnh: {e}")

    if st.session_state.parsed_target is not None and st.session_state.dashboard_results:
        groups = st.session_state.parsed_target["Nhóm hàng"].astype(str).tolist()
        summary = []
        for code, name in EMPLOYEES:
            r = st.session_state.dashboard_results[code]
            active = sum(1 for c in r["cards"] if c["Nhóm ghép"] and c["Tỷ lệ (%)"] is not None)
            summary.append({
                "Mã NV": code, "Nhân viên": name, "Số thẻ phát hiện": r["card_count"],
                "Thẻ ghép được": active, "Số 0 trong danh sách": r["zero_matched"],
                "Số 0 theo header": r["zero_expected"],
                "Ô đã nhận diện": len(r["values"]), "Ô cần rà soát": len(groups) - len(r["values"])
            })
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)
        st.caption("Số ô đã nhận diện gồm tỷ lệ từ thẻ và số 0 chỉ khi có tên trong danh sách 'Chưa có số liệu'.")
        with st.expander("Xem nhật ký từng thẻ OCR"):
            st.dataframe(st.session_state.card_audit, use_container_width=True, hide_index=True)
        with st.expander("Xem nguyên văn danh sách 'Chưa có số liệu'"):
            for code, name in EMPLOYEES:
                st.markdown(f"**{name} — {code}**")
                st.text(st.session_state.dashboard_results[code]["zero_raw"][:2500])

with tabs[2]:
    if st.session_state.parsed_target is None or st.session_state.dashboard_results is None:
        st.info("Chạy phân tích ảnh ở tab ② trước.")
    else:
        st.markdown("### A. Rà soát danh mục, THỰC HIỆN và TARGET từ ảnh Lũy kế")
        st.caption("Anh có thể sửa tên nhóm, THỰC HIỆN hoặc TARGET nếu OCR đọc chưa đúng. Thứ tự dòng được giữ nguyên.")
        target_input = st.data_editor(
            st.session_state.target_editor,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="target_table_v5",
            column_config={
                "STT": st.column_config.NumberColumn("STT", disabled=True),
                "Nhóm hàng": st.column_config.TextColumn("Tên nhóm hàng", required=True),
                "THỰC HIỆN": st.column_config.NumberColumn("THỰC HIỆN cửa hàng", format="%.2f", required=True),
                "TARGET siêu thị": st.column_config.NumberColumn("TARGET siêu thị", format="%.2f", required=False),
                "OCR tên": st.column_config.TextColumn("OCR gốc", disabled=True),
                "OCR số": st.column_config.TextColumn("OCR số gốc", disabled=True),
                "Trạng thái": st.column_config.TextColumn(disabled=True)
            }
        )
        st.session_state.target_editor = target_input
        group_names = target_input["Nhóm hàng"].astype(str).tolist()
        if len(set(norm(x) for x in group_names)) != len(group_names):
            st.error("Tên nhóm hàng bị trùng sau khi chuẩn hóa. Hãy sửa tên trước khi tạo ma trận.")
        if target_input["THỰC HIỆN"].isna().any() or not all(str(x).strip() for x in group_names):
            st.warning("Còn tên nhóm hoặc TARGET chưa đọc được. Hãy sửa các ô trống trước khi tạo ma trận.")

        st.divider()
        st.markdown("### B. Tạo ma trận tỷ lệ 9 nhân viên × nhóm hàng")
        st.caption("Tỷ lệ từ thẻ được giữ nguyên, 0 chỉ lấy từ danh sách 'Chưa có số liệu'. Ô không đọc được để trống.")
        if st.button("TẠO / CẬP NHẬT MA TRẬN", type="primary"):
            if target_input["THỰC HIỆN"].isna().any() or not all(str(x).strip() for x in group_names):
                st.error("Hãy hoàn tất tên nhóm và THỰC HIỆN trước. TARGET có thể để trống nếu ảnh ghi không có target.")
            elif len(set(norm(x) for x in group_names)) != len(group_names):
                st.error("Tên nhóm hàng bị trùng.")
            else:
                rows = []
                for code, name in EMPLOYEES:
                    res = st.session_state.dashboard_results[code]
                    row = {"Nhân viên": f"{name} - {code}"}
                    for i, group in enumerate(group_names):
                        row[group] = res["values"].get(i, None)
                    rows.append(row)
                st.session_state.matrix_editor = pd.DataFrame(rows)
                st.session_state.report_matrix = None
                st.success(f"Đã tạo ma trận {len(rows)} nhân viên × {len(group_names)} nhóm hàng.")

        if st.session_state.matrix_editor is not None:
            st.markdown("### C. Sửa các ô OCR chưa chắc chắn")
            matrix_input = st.data_editor(
                st.session_state.matrix_editor,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                height=560,
                key="matrix_v5",
                column_config={
                    "Nhân viên": st.column_config.TextColumn(disabled=True),
                    **{g: st.column_config.NumberColumn(g, format="%.2f") for g in group_names}
                }
            )
            st.session_state.matrix_editor = matrix_input
            missing = int(matrix_input[group_names].isna().sum().sum())
            st.metric("Ô tỷ lệ còn trống / chưa xác minh", missing)
            if missing:
                st.warning("Còn ô trống. Anh xác nhận từ ảnh gốc rồi nhập tỷ lệ hoặc nhập 0 nếu dashboard ghi rõ không có số liệu.")
            else:
                st.success("Đã đủ toàn bộ tỷ lệ cho ma trận. Chuyển sang tab ④ để xem trước doanh số.")

        if st.session_state.card_audit is not None:
            with st.expander("Các thẻ cần kiểm tra tên nhóm / tỷ lệ"):
                audit = st.session_state.card_audit
                needs = audit[audit["Trạng thái"] != "Đã ghép"]
                if len(needs):
                    st.dataframe(needs, use_container_width=True, hide_index=True)
                else:
                    st.success("Tất cả thẻ OCR đã ghép được tên nhóm và tỷ lệ.")

with tabs[3]:
    if st.session_state.matrix_editor is None or st.session_state.target_editor is None:
        st.info("Hoàn tất bước ③ để xem trước.")
    else:
        groups = st.session_state.target_editor["Nhóm hàng"].astype(str).tolist()
        targets = pd.to_numeric(st.session_state.target_editor["TARGET siêu thị"], errors="coerce")
        actuals = pd.to_numeric(st.session_state.target_editor["THỰC HIỆN"], errors="coerce")
        rates = st.session_state.matrix_editor.copy()
        missing_rates = int(rates[groups].isna().sum().sum())
        missing_targets = int(actuals.isna().sum())
        if missing_rates or missing_targets:
            st.warning(f"Còn {missing_rates} ô tỷ lệ và {missing_targets} TARGET chưa xác minh. Bảng sẽ giữ trống các ô chưa đủ dữ liệu.")
        sales = rates.copy()
        for j, group in enumerate(groups):
            weights = pd.to_numeric(rates[group], errors="coerce")
            weight_total = weights.sum(min_count=staff_count)
            if pd.isna(actuals.iloc[j]) or pd.isna(weight_total):
                sales[group] = np.nan
            elif weight_total <= 0 and abs(float(actuals.iloc[j])) < 1e-12:
                sales[group] = 0.0
            elif weight_total <= 0:
                sales[group] = np.nan
            else:
                sales[group] = actuals.iloc[j] * weights / weight_total
        st.markdown("### Bảng doanh số phân bổ — tổng 9 nhân viên khớp THỰC HIỆN cửa hàng")
        preview = sales.copy()
        total = {"Nhân viên": "TỔNG"}
        for group in groups:
            col = pd.to_numeric(sales[group], errors="coerce")
            total[group] = float(col.sum()) if col.notna().all() else None
        preview = pd.concat([pd.DataFrame([total]), preview], ignore_index=True)

        def row_style(row):
            bg = "#FFC000" if row.name == 0 else "#DDEBF7"
            weight = "bold" if row.name == 0 else "normal"
            return [f"background-color:{bg};color:#000;border:1px solid #202020;font-weight:{weight}"] * len(row)

        styled = (preview.style.apply(row_style, axis=1)
            .set_table_styles([
                {"selector":"th","props":[("background-color","#FFF200"),("color","#000"),("font-weight","bold"),("border","1px solid #202020"),("white-space","normal")]},
                {"selector":"td","props":[("border","1px solid #202020"),("text-align","center"),("padding","4px")]},
                {"selector":"th:first-child,td:first-child","props":[("min-width","190px"),("text-align","left")]}
            ])
            .format({g:"{:,.2f}" for g in groups}, na_rep=""))
        st.dataframe(styled, use_container_width=True, height=540, hide_index=True)

        # Reconciliation check: sum of employee allocations must equal store actual for each group.
        recon = []
        for j, group in enumerate(groups):
            ssum = pd.to_numeric(sales[group], errors="coerce").sum(min_count=staff_count)
            store_actual = actuals.iloc[j]
            recon.append({"Nhóm hàng": group, "THỰC HIỆN cửa hàng": store_actual,
                          "Tổng 9 NV": ssum,
                          "Chênh lệch": (ssum - store_actual) if pd.notna(ssum) and pd.notna(store_actual) else np.nan})
        recon_df = pd.DataFrame(recon)
        with st.expander("Đối soát tổng doanh số 9 NV với THỰC HIỆN cửa hàng"):
            st.dataframe(recon_df, use_container_width=True, hide_index=True)
            diffs = pd.to_numeric(recon_df["Chênh lệch"], errors="coerce").dropna().abs()
            if len(diffs) == len(groups) and (diffs < 0.000001).all():
                st.success("Đã đối soát: tổng doanh số 9 NV khớp THỰC HIỆN cửa hàng ở tất cả nhóm.")
            else:
                st.warning("Còn nhóm chưa đủ tỷ lệ hoặc chưa khớp tổng THỰC HIỆN. Chưa nên chốt báo cáo.")

        st.markdown("### Tình trạng xác minh")
        c1, c2, c3 = st.columns(3)
        c1.metric("Nhóm hàng tháng 11", len(groups))
        c2.metric("Nhân viên", len(sales))
        c3.metric("Ô tỷ lệ cần rà soát", missing_rates)

        audit_df = st.session_state.card_audit if st.session_state.card_audit is not None else pd.DataFrame()
        # Output as draft when blanks remain; blank cells stay blank, never silently become zero.
        xlsx = build_excel(groups, sales, rates, report_date, staff_count, audit_df)
        png = build_png(groups, sales, report_date)
        a, b = st.columns(2)
        a.download_button(
            "⬇ TẢI FILE EXCEL", xlsx,
            f"Bao_cao_Thang_11_{report_date.replace('/','-')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        b.download_button(
            "⬇ TẢI ẢNH PNG", png,
            f"Bao_cao_Thang_11_{report_date.replace('/','-')}.png",
            "image/png",
            use_container_width=True
        )
        if missing_rates or missing_targets:
            st.warning("Đây là bản nháp: các ô chưa xác minh đang để trống. Không dùng để chốt thi đua cho đến khi hoàn tất.")
        else:
            st.success("Đã đủ dữ liệu đầu vào. Vui lòng đối chiếu ngẫu nhiên với ảnh gốc trước khi chốt.")
