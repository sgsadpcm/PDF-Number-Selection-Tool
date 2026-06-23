import os
import hashlib
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "regulation-lookup-secret"

BASE_DIR      = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
STATIC_IMG    = os.path.join(BASE_DIR, "static", "images")
ALLOWED_EXCEL = {"xlsx", "xls"}
ALLOWED_IMG   = {"png", "jpg", "jpeg", "gif", "webp"}
EXCEL_FILE    = os.path.join(UPLOAD_FOLDER, "data.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_IMG, exist_ok=True)

COL_DEFECT  = "缺失項目"
COL_REG     = "法源依據"
COL_CONTENT = "法規內容"

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_SS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# ─── helpers ─────────────────────────────────────────────────────────────────

def allowed_excel(f):
    return "." in f and f.rsplit(".", 1)[1].lower() in ALLOWED_EXCEL

def allowed_img(f):
    return "." in f and f.rsplit(".", 1)[1].lower() in ALLOWED_IMG

def get_excel_path():
    if os.path.exists(EXCEL_FILE):
        return EXCEL_FILE
    files = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith((".xlsx", ".xls"))]
    return os.path.join(UPLOAD_FOLDER, files[0]) if files else None

def find_col(columns, target):
    t = target.strip()
    for c in columns:
        if str(c).strip() == t:
            return c
    return None


# ─── xlsx image extraction ────────────────────────────────────────────────────

def _build_sheet_drawing_map(z):
    """
    回傳 { sheet_file_number: drawing_number } 以及 { sheet_name: sheet_file_number }
    """
    # workbook.xml → sheet name → rId
    wb_xml   = ET.fromstring(z.read("xl/workbook.xml").decode())
    wb_rels  = ET.fromstring(z.read("xl/_rels/workbook.xml.rels").decode())

    # rId → sheet file number
    rid_to_sheetnum = {}
    for rel in wb_rels:
        target = rel.get("Target", "")
        rid    = rel.get("Id", "")
        if "sheet" in target.lower():
            num = target.replace("worksheets/sheet", "").replace(".xml", "")
            rid_to_sheetnum[rid] = num

    # sheet name → sheet file number
    name_to_sheetnum = {}
    ns = {"ns": NS_SS, "r": NS_R}
    for s in wb_xml.findall(".//ns:sheet", ns) or wb_xml.findall(".//sheet"):
        name = s.get("name", "")
        rid  = s.get(f"{{{NS_R}}}id", "") or s.get("r:id", "")
        if rid in rid_to_sheetnum:
            name_to_sheetnum[name] = rid_to_sheetnum[rid]

    # sheet file number → drawing number
    sheetnum_to_drawing = {}
    all_files = z.namelist()
    for sheetnum in rid_to_sheetnum.values():
        rels_path = f"xl/worksheets/_rels/sheet{sheetnum}.xml.rels"
        if rels_path not in all_files:
            continue
        rels_root = ET.fromstring(z.read(rels_path).decode())
        for rel in rels_root:
            target = rel.get("Target", "")
            rtype  = rel.get("Type", "")
            if "drawing" in rtype.lower() or "drawing" in target.lower():
                # "../drawings/drawing3.xml" → 3
                d_num = target.split("drawing")[-1].replace(".xml", "")
                sheetnum_to_drawing[sheetnum] = d_num

    return name_to_sheetnum, sheetnum_to_drawing


def _parse_drawing_images(z, d_num):
    """
    回傳 [ {row: int, col: int, media_path: str} ]
    row 是 0-indexed，對應 Excel 行號（0 = 第一列 = header）
    """
    d_path    = f"xl/drawings/drawing{d_num}.xml"
    d_rels    = f"xl/drawings/_rels/drawing{d_num}.xml.rels"
    all_files = z.namelist()
    if d_path not in all_files:
        return []

    # rId → media path
    rid_to_media = {}
    if d_rels in all_files:
        rels_root = ET.fromstring(z.read(d_rels).decode())
        for rel in rels_root:
            rid    = rel.get("Id", "")
            target = rel.get("Target", "")   # "../media/imageN.png"
            media  = "xl/" + target.replace("../", "")
            rid_to_media[rid] = media

    result = []
    root = ET.fromstring(z.read(d_path).decode())
    for anchor in root:
        tag = anchor.tag.split("}")[-1]
        if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        from_el = anchor.find(f"{{{NS_XDR}}}from")
        row = col = 0
        if from_el is not None:
            r_el = from_el.find(f"{{{NS_XDR}}}row")
            c_el = from_el.find(f"{{{NS_XDR}}}col")
            row = int(r_el.text) if r_el is not None else 0
            col = int(c_el.text) if c_el is not None else 0
        blip = anchor.find(f".//{{{NS_A}}}blip")
        if blip is None:
            continue
        rid = blip.get(f"{{{NS_R}}}embed", "")
        if rid in rid_to_media:
            result.append({"row": row, "col": col, "media": rid_to_media[rid]})
    return result


def extract_and_cache_images(sheet_name, defect_row_ranges):
    """
    從 xlsx 提取圖片，依 drawing row 對應到 defect index，
    存到 static/images/<sheet>/<defect_idx>/，
    回傳 { defect_idx: [filename, ...] }
    已存在的圖片不重複提取。
    """
    path = get_excel_path()
    if not path:
        return {}
    try:
        with zipfile.ZipFile(path) as z:
            name_to_sheetnum, sheetnum_to_drawing = _build_sheet_drawing_map(z)

            sheetnum = name_to_sheetnum.get(sheet_name)
            if sheetnum is None:
                return {}
            d_num = sheetnum_to_drawing.get(sheetnum)
            if d_num is None:
                return {}

            drawing_imgs = _parse_drawing_images(z, d_num)
            result = {}

            for img_info in drawing_imgs:
                draw_row  = img_info["row"]
                media_key = img_info["media"]

                # 找對應的 defect index
                defect_idx = None
                for i, (start_row, end_row) in enumerate(defect_row_ranges):
                    if start_row <= draw_row <= end_row:
                        defect_idx = i
                        break
                if defect_idx is None:
                    continue

                if media_key not in z.namelist():
                    continue
                img_bytes = z.read(media_key)
                ext = media_key.rsplit(".", 1)[-1].lower()

                # EMF → PNG via ImageMagick
                if ext == "emf":
                    ext = "png"
                    img_bytes = _convert_emf_to_png(img_bytes)
                    if img_bytes is None:
                        continue

                img_hash = hashlib.md5(img_bytes).hexdigest()[:10]
                filename = f"exc_{img_hash}.{ext}"
                folder   = _img_folder(sheet_name, defect_idx)
                filepath = os.path.join(folder, filename)
                if not os.path.exists(filepath):
                    with open(filepath, "wb") as f:
                        f.write(img_bytes)
                result.setdefault(defect_idx, []).append(filename)
    except Exception as e:
        app.logger.error(f"extract_and_cache_images error: {e}")
        return {}
    return result


def _convert_emf_to_png(emf_bytes):
    """用 ImageMagick magick 把 EMF 轉成 PNG bytes，失敗回傳 None"""
    import tempfile
    try:
        with tempfile.NamedTemporaryFile(suffix=".emf", delete=False) as f_in:
            f_in.write(emf_bytes)
            in_path = f_in.name
        out_path = in_path.replace(".emf", ".png")
        result = subprocess.run(
            ["magick", in_path, out_path],
            capture_output=True, timeout=15
        )
        if result.returncode == 0 and os.path.exists(out_path):
            with open(out_path, "rb") as f:
                data = f.read()
            return data
        return None
    except Exception:
        return None
    finally:
        try:
            os.unlink(in_path)
        except Exception:
            pass
        try:
            os.unlink(out_path)
        except Exception:
            pass


def _sheet_dir(sheet_name):
    """用 hash 當資料夾名稱，避免中文被 secure_filename 全部清空而混用同一目錄"""
    return hashlib.md5(sheet_name.encode("utf-8")).hexdigest()[:12]


def _img_folder(sheet_name, item_index):
    folder = os.path.join(STATIC_IMG, _sheet_dir(sheet_name), str(item_index))
    os.makedirs(folder, exist_ok=True)
    return folder


def list_images(sheet_name, item_index):
    folder = _img_folder(sheet_name, item_index)
    return sorted([f for f in os.listdir(folder) if f.rsplit(".", 1)[-1].lower() in ALLOWED_IMG])


# ─── data loading ─────────────────────────────────────────────────────────────

def load_sheets():
    path = get_excel_path()
    if not path:
        return None, None
    try:
        return pd.ExcelFile(path).sheet_names, path
    except Exception as e:
        return None, str(e)


def load_sheet_data(sheet_name):
    path = get_excel_path()
    if not path:
        return None, [], []

    try:
        # 第一次讀（自動判斷）
        raw = pd.read_excel(path, sheet_name=sheet_name)
        raw = raw.fillna("")
        raw.columns = [str(c).strip() for c in raw.columns]
        raw = raw.dropna(how="all")

        actual_cols = list(raw.columns)

        # 找欄位
        col_defect = find_col(raw.columns, COL_DEFECT)
        col_reg = find_col(raw.columns, COL_REG)

        # 如果抓不到 → 換 header=1
        if col_defect is None:
            raw = pd.read_excel(path, sheet_name=sheet_name, header=1)
            raw = raw.fillna("")
            raw.columns = [str(c).strip() for c in raw.columns]
            raw = raw.dropna(how="all")

            actual_cols = list(raw.columns)
            col_defect = find_col(raw.columns, COL_DEFECT)
            col_reg = find_col(raw.columns, COL_REG)

        if col_defect is None:
            return None, actual_cols, []

        records = []
        row_ranges = []
        current = None
        current_start = None

        for enum_idx, (_, row) in enumerate(raw.iterrows()):
            draw_row = enum_idx + 1

            val = str(row[col_defect]).strip() if pd.notna(row[col_defect]) else ""
            reg_val = str(row[col_reg]).strip() if col_reg and pd.notna(row[col_reg]) else ""

            if val:
                if current is not None:
                    row_ranges.append((current_start, draw_row - 1))
                    records.append(current)

                current = {
                    COL_DEFECT: val,
                    COL_REG: reg_val,
                    COL_CONTENT: ""
                }
                current_start = draw_row
            else:
                if current is not None and reg_val:
                    if current[COL_CONTENT]:
                        current[COL_CONTENT] += "\n"
                    current[COL_CONTENT] += reg_val

        if current is not None:
            row_ranges.append((current_start, 99999))
            records.append(current)

        df = pd.DataFrame(records) if records else pd.DataFrame(
            columns=[COL_DEFECT, COL_REG, COL_CONTENT]
        )

        return df, actual_cols, row_ranges

    except Exception as e:
        return None, [str(e)], []
# ─── routes ───────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if not f or f.filename == "":
            flash("請選擇檔案")
        elif allowed_excel(f.filename):
            for old in os.listdir(UPLOAD_FOLDER):
                os.remove(os.path.join(UPLOAD_FOLDER, old))
            # 清除圖片快取
            import shutil
            if os.path.exists(STATIC_IMG):
                shutil.rmtree(STATIC_IMG)
            os.makedirs(STATIC_IMG, exist_ok=True)
            f.save(EXCEL_FILE)
            flash("檔案上傳成功！")
        else:
            flash("僅支援 .xlsx 或 .xls 格式")
        return redirect(url_for("index"))

    sheets, error = load_sheets()
    return render_template("index.html", sheets=sheets, error=error)


@app.route("/system/<path:sheet_name>")
def defects(sheet_name):
    df, _, _ = load_sheet_data(sheet_name)
    if df is None:
        return redirect(url_for("index"))

    # ===== Excel資料 =====
    excel_items = df[COL_DEFECT].tolist()

    # ===== JSON新增資料 =====
    data = load_json()
    extra = data.get(sheet_name, [])

    # ===== 合併（新增放前面）=====
    items = []

    for item in extra:
        items.append("🆕 " + item["缺失項目"])

    items += excel_items

    return render_template("defects.html", sheet_name=sheet_name, items=items)
    df, _, _ = load_sheet_data(sheet_name)
    if df is None:
        return redirect(url_for("index"))
    items = df[COL_DEFECT].tolist()
    return render_template("defects.html", sheet_name=sheet_name, items=items)


@app.route("/system/<path:sheet_name>/defect/<int:item_index>", methods=["GET", "POST"])
def regulation(sheet_name, item_index):
    df, actual_cols, row_ranges = load_sheet_data(sheet_name)
    if df is None or item_index >= len(df):
        return redirect(url_for("index"))

    # 手動上傳圖片
    if request.method == "POST":
        uploaded = request.files.getlist("images")
        count = 0
        for f in uploaded:
            if f and f.filename and allowed_img(f.filename):
                ext  = f.filename.rsplit(".", 1)[1].lower()
                name = hashlib.md5(f.read()).hexdigest()[:12] + "." + ext
                f.seek(0)
                f.save(os.path.join(_img_folder(sheet_name, item_index), name))
                count += 1
        if count:
            flash(f"上傳了 {count} 張圖片")
        return redirect(url_for("regulation", sheet_name=sheet_name, item_index=item_index))

    # 從 Excel 提取圖片（只提取這個工作表一次）
    extract_and_cache_images(sheet_name, row_ranges)

    row          = df.iloc[item_index]
    defect       = row[COL_DEFECT]
    reg_text     = row[COL_REG]
    content_text = row[COL_CONTENT]
    images       = list_images(sheet_name, item_index)
    col_warning  = actual_cols if (not reg_text and not content_text and not images) else None

    return render_template(
        "regulation.html",
        sheet_name=sheet_name,
        defect=defect,
        reg_text=reg_text,
        content_text=content_text,
        images=images,
        item_index=item_index,
        col_warning=col_warning,
    )


@app.route("/system/<path:sheet_name>/defect/<int:item_index>/delete_image/<filename>")
def delete_image(sheet_name, item_index, filename):
    try:
        os.remove(os.path.join(_img_folder(sheet_name, item_index), secure_filename(filename)))
    except Exception:
        pass
    return redirect(url_for("regulation", sheet_name=sheet_name, item_index=item_index))


@app.route("/img/<path:sheet_name>/<int:item_index>/<filename>")
def serve_image(sheet_name, item_index, filename):
    return send_from_directory(_img_folder(sheet_name, item_index), filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
# ================== JSON 資料（新增缺失用） ==================
import json

DATA_JSON = os.path.join(BASE_DIR, "data", "defects.json")

def load_json():
    if not os.path.exists(DATA_JSON):
        return {}
    with open(DATA_JSON, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json(data):
    os.makedirs(os.path.dirname(DATA_JSON), exist_ok=True)
    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ================== 新增缺失頁面 ==================
@app.route("/add_defect/<sheet_name>", methods=["GET", "POST"])
def add_defect(sheet_name):
    if request.method == "POST":
        defect = request.form.get("defect")
        reg = request.form.get("reg")

        if not defect:
            return redirect(url_for("defects", sheet_name=sheet_name))

        data = load_json()
        data.setdefault(sheet_name, []).append({
            "缺失項目": defect,
            "法源依據": reg
        })
        save_json(data)

        return redirect(url_for("defects", sheet_name=sheet_name))

    return render_template("add_defect.html", sheet_name=sheet_name)