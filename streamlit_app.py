import io
import json
import mimetypes
import os
import re
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from string import Formatter

import fitz
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


FNSKU_REGEX = r"\bX[A-Z0-9]{9}\b"
FILENAME_TEMPLATE = "{SKU}-{FNSKU}.pdf"
LABEL_TEXT_LINE2_TEMPLATE = "{款式}"

PDF_TEXT_COVER_TOP_RATIO = 0.535
PDF_FNSKU_BASELINE_RATIO = 0.660
PDF_LINE1_BASELINE_RATIO = 0.783
PDF_LINE2_BASELINE_RATIO = 0.895
PDF_SIDE_MARGIN = 4

CJK_FONT_CANDIDATES = []
CJK_BOLD_STROKE_WIDTH = 0.05

FEISHU_INDEX_RANGE_DEFAULT = "A1:G50000"
SKIP_SHEET_NAMES = {"更新日志", "更新记录", "日志", "说明", "README"}

LOCAL_ILLEGAL_FILENAME_CHARS = r'\/:*?"<>|'

LOG_HEADERS = [
    "处理时间",
    "原文件名",
    "页码",
    "识别FNSKU",
    "匹配状态",
    "失败原因",
    "标签第一行",
    "标签第二行",
    "飞书正式文件名",
    "本地备份文件名",
    "飞书文件token",
    "处理动作",
]

REQUIRED_INDEX_HEADERS = ["FNSKU", "SKU", "款式", "品牌", "型号", "产品类型"]
REQUIRED_INFO_FIELDS = ["SKU", "款式", "品牌", "型号", "产品类型"]
OPTIONAL_INFO_FIELDS = ["颜色"]


class FatalError(Exception):
    pass


class FileNameError(Exception):
    pass


class FeishuError(Exception):
    pass


def cell_to_text(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_fnsku(value):
    return cell_to_text(value).upper()


def is_valid_fnsku(value):
    return re.fullmatch(FNSKU_REGEX, value, flags=re.IGNORECASE) is not None


def contains_cjk(text):
    return any("\u4e00" <= char <= "\u9fff" for char in text)


def is_cjk_char(char):
    return "\u4e00" <= char <= "\u9fff"


def find_cjk_fontfile():
    for font_path in CJK_FONT_CANDIDATES:
        if Path(font_path).exists():
            return font_path
    return ""


def ensure_pdf_suffix(filename):
    if filename.lower().endswith(".pdf"):
        return filename[:-4] + ".pdf"
    return filename + ".pdf"


def normalize_display_filename(filename):
    filename = cell_to_text(filename)
    filename = re.sub(r"[\r\n\t]+", " ", filename)
    filename = re.sub(r"\s+", "-", filename)
    filename = filename.strip(" .-")
    if not filename:
        raise FileNameError("文件名为空")

    filename = ensure_pdf_suffix(filename)
    if filename.lower() == ".pdf":
        raise FileNameError("文件名为空")
    return filename


def make_safe_local_filename(display_filename, fnsku):
    filename = normalize_display_filename(display_filename)
    for char in LOCAL_ILLEGAL_FILENAME_CHARS:
        filename = filename.replace(char, "-")
    filename = re.sub(r"-+", "-", filename)
    filename = filename.strip(" .-")
    if not filename:
        filename = f"{fnsku}.pdf"
    return ensure_pdf_suffix(filename)


def get_template_fields(template):
    fields = []
    for _, field_name, _, _ in Formatter().parse(template):
        if not field_name:
            continue
        simple_name = field_name.split(".", 1)[0].split("[", 1)[0]
        if simple_name and simple_name not in fields:
            fields.append(simple_name)
    return fields


def build_display_filename(row, fnsku):
    values = {}
    missing_fields = []

    for field in get_template_fields(FILENAME_TEMPLATE):
        if field == "FNSKU":
            value = fnsku
        else:
            value = cell_to_text(row.get(field))

        if not value:
            missing_fields.append(field)
        values[field] = value

    if missing_fields:
        raise FileNameError("生成文件名所需信息缺失：" + "、".join(missing_fields))

    try:
        filename = normalize_display_filename(FILENAME_TEMPLATE.format(**values))
    except KeyError as exc:
        raise FileNameError(f"文件名模板字段不存在：{exc}") from exc

    if fnsku.upper() not in filename.upper():
        raise FileNameError("输出文件名未包含识别到的 FNSKU")
    return filename


def build_label_lines(row):
    missing_fields = [field for field in REQUIRED_INFO_FIELDS if not cell_to_text(row.get(field))]
    if missing_fields:
        raise FileNameError("标签信息缺失：" + "、".join(missing_fields))

    brand = cell_to_text(row.get("品牌"))
    model = cell_to_text(row.get("型号"))
    product_type = cell_to_text(row.get("产品类型"))
    color = cell_to_text(row.get("颜色"))

    line1 = f"{brand} for {model} {product_type}"
    if color:
        line1 = f"{line1}, {color}"

    try:
        line2 = LABEL_TEXT_LINE2_TEMPLATE.format(**row)
    except KeyError as exc:
        raise FileNameError(f"标签文字模板字段不存在：{exc}") from exc

    line1 = cell_to_text(line1)
    line2 = cell_to_text(line2)
    if not line1 or not line2:
        raise FileNameError("标签文字为空")
    return line1, line2


def get_draw_font(text, cjk_fontfile):
    if contains_cjk(text):
        if cjk_fontfile:
            return "labelcjk", cjk_fontfile, fitz.Font(fontfile=cjk_fontfile)
        return "china-ss", "", fitz.Font(fontname="china-ss")
    return "helv", "", fitz.Font(fontname="helv")


def split_font_runs(text):
    runs = []
    current_text = ""
    current_is_cjk = None

    for char in text:
        char_is_cjk = is_cjk_char(char)
        if current_text and char_is_cjk != current_is_cjk:
            runs.append((current_text, current_is_cjk))
            current_text = char
        else:
            current_text += char
        current_is_cjk = char_is_cjk

    if current_text:
        runs.append((current_text, current_is_cjk))
    return runs


def measure_runs_width(runs, font_size, cjk_fontfile):
    total_width = 0
    prepared_runs = []

    for run_text, _ in runs:
        fontname, fontfile, font = get_draw_font(run_text, cjk_fontfile)
        width = font.text_length(run_text, fontsize=font_size)
        prepared_runs.append((run_text, fontname, fontfile, width))
        total_width += width

    return total_width, prepared_runs


def fit_runs_font_size(runs, cjk_fontfile, max_width, start_size, min_size):
    font_size = start_size
    while font_size > min_size:
        total_width, _ = measure_runs_width(runs, font_size, cjk_fontfile)
        if total_width <= max_width:
            break
        font_size -= 0.2
    return max(font_size, min_size)


def draw_centered_text(page, text, baseline_y, start_size, min_size, cjk_fontfile):
    text = cell_to_text(text)
    if not text:
        return

    runs = split_font_runs(text)
    max_width = page.rect.width - PDF_SIDE_MARGIN * 2 - 2
    font_size = fit_runs_font_size(runs, cjk_fontfile, max_width, start_size, min_size)
    text_width, prepared_runs = measure_runs_width(runs, font_size, cjk_fontfile)
    x = (page.rect.width - text_width) / 2

    for run_text, fontname, fontfile, width in prepared_runs:
        kwargs = {
            "fontsize": font_size,
            "fontname": fontname,
            "color": (0, 0, 0),
            "fill": (0, 0, 0),
            "overlay": True,
        }
        if contains_cjk(run_text) and CJK_BOLD_STROKE_WIDTH > 0:
            kwargs["render_mode"] = 2
            kwargs["border_width"] = CJK_BOLD_STROKE_WIDTH
        if fontfile:
            kwargs["fontfile"] = fontfile

        page.insert_text(fitz.Point(x, baseline_y), run_text, **kwargs)
        x += width


def rewrite_label_pdf(source_path, target_path, page_index, fnsku, label_line1, label_line2):
    target_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = target_path.parent / f".{target_path.stem}.{uuid.uuid4().hex}.tmp.pdf"
    cjk_fontfile = find_cjk_fontfile()

    try:
        with fitz.open(source_path) as source_document:
            if source_document.page_count == 0:
                raise RuntimeError("PDF 没有页面")
            if page_index < 0 or page_index >= source_document.page_count:
                raise RuntimeError(f"PDF 页码不存在：第 {page_index + 1} 页")

            document = fitz.open()
            try:
                document.insert_pdf(source_document, from_page=page_index, to_page=page_index)
                page = document[0]
                rect = page.rect
                scale = rect.width / 142.08 if rect.width else 1

                erase_rect = fitz.Rect(
                    PDF_SIDE_MARGIN,
                    rect.height * PDF_TEXT_COVER_TOP_RATIO,
                    rect.width - PDF_SIDE_MARGIN,
                    rect.height - 2,
                )
                page.add_redact_annot(erase_rect, fill=(1, 1, 1))
                page.apply_redactions()

                draw_centered_text(
                    page,
                    fnsku,
                    rect.height * PDF_FNSKU_BASELINE_RATIO,
                    10.5 * scale,
                    7.5 * scale,
                    cjk_fontfile,
                )
                draw_centered_text(
                    page,
                    label_line1,
                    rect.height * PDF_LINE1_BASELINE_RATIO,
                    6.9 * scale,
                    4.8 * scale,
                    cjk_fontfile,
                )
                draw_centered_text(
                    page,
                    label_line2,
                    rect.height * PDF_LINE2_BASELINE_RATIO,
                    7.4 * scale,
                    5.0 * scale,
                    cjk_fontfile,
                )

                document.save(temp_path, garbage=4, deflate=True)
            finally:
                document.close()

        os.replace(temp_path, target_path)
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise

    return "生成修改后PDF"


def format_sheet_row(sheet_name, row_number):
    return f"「{sheet_name}」第 {row_number} 行"


def sheet_should_skip(sheet_name):
    return cell_to_text(sheet_name) in SKIP_SHEET_NAMES


def build_header_to_index(values):
    if not values:
        return {}
    header_values = [cell_to_text(value) for value in values[0]]
    return {header: idx for idx, header in enumerate(header_values) if header}


def sheet_looks_like_index(header_to_index):
    return any(header in header_to_index for header in REQUIRED_INDEX_HEADERS)


def parse_index_sheets(sheet_values_list):
    missing_fnsku_rows = []
    invalid_fnsku_rows = []
    skipped_same_duplicate_rows = []
    skipped_sheet_names = []
    used_sheet_names = []
    index_rows = {}
    duplicate_conflicts = defaultdict(list)
    duplicate_check_fields = REQUIRED_INFO_FIELDS + OPTIONAL_INFO_FIELDS

    for sheet_name, values in sheet_values_list:
        sheet_name = cell_to_text(sheet_name) or "未命名Sheet"
        if sheet_should_skip(sheet_name):
            skipped_sheet_names.append(sheet_name)
            continue

        header_to_index = build_header_to_index(values)
        if not sheet_looks_like_index(header_to_index):
            skipped_sheet_names.append(sheet_name)
            continue

        missing_headers = [header for header in REQUIRED_INDEX_HEADERS if header not in header_to_index]
        if missing_headers:
            raise FatalError(f"工作表「{sheet_name}」像索引表，但缺少表头：" + "、".join(missing_headers))

        used_sheet_names.append(sheet_name)

        for row_number, row_values in enumerate(values[1:], start=2):
            row = {}
            has_any_value = False

            for header, idx in header_to_index.items():
                value = row_values[idx] if idx < len(row_values) else None
                text = cell_to_text(value)
                row[header] = text
                if text:
                    has_any_value = True

            if not has_any_value:
                continue

            row_ref = format_sheet_row(sheet_name, row_number)
            fnsku = normalize_fnsku(row.get("FNSKU"))
            if not fnsku:
                missing_fnsku_rows.append(row_ref)
                continue

            if not is_valid_fnsku(fnsku):
                invalid_fnsku_rows.append((row_ref, fnsku))
                continue

            row["FNSKU"] = fnsku
            row["_索引位置"] = row_ref

            existing_row = index_rows.get(fnsku)
            if existing_row:
                is_same_content = all(
                    cell_to_text(existing_row.get(field)) == cell_to_text(row.get(field))
                    for field in duplicate_check_fields
                )
                if is_same_content:
                    skipped_same_duplicate_rows.append((row_ref, fnsku, existing_row["_索引位置"]))
                    continue

                if not duplicate_conflicts[fnsku]:
                    duplicate_conflicts[fnsku].append(existing_row["_索引位置"])
                duplicate_conflicts[fnsku].append(row_ref)
                continue

            index_rows[fnsku] = row

    if not used_sheet_names:
        raise FatalError("没有找到索引 Sheet：请确认至少一个 Sheet 第一行包含 FNSKU、SKU 等表头")

    if missing_fnsku_rows:
        raise FatalError(
            "索引表中以下位置缺少必填 FNSKU："
            + "、".join(missing_fnsku_rows[:30])
            + ("、..." if len(missing_fnsku_rows) > 30 else "")
        )

    warnings = []
    if invalid_fnsku_rows:
        sample_text = "；".join(f"{row_ref} {fnsku}" for row_ref, fnsku in invalid_fnsku_rows[:10])
        if len(invalid_fnsku_rows) > 10:
            sample_text += "；..."
        warnings.append(f"已忽略 {len(invalid_fnsku_rows)} 行无效 FNSKU（{sample_text}）")

    if skipped_same_duplicate_rows:
        sample_text = "；".join(
            f"{fnsku} {row_ref}（与 {first_row_ref} 一致）"
            for row_ref, fnsku, first_row_ref in skipped_same_duplicate_rows[:10]
        )
        if len(skipped_same_duplicate_rows) > 10:
            sample_text += "；..."
        warnings.append(f"已忽略 {len(skipped_same_duplicate_rows)} 行完全相同的重复 FNSKU（{sample_text}）")

    if duplicate_conflicts:
        lines = ["索引表中 FNSKU 重复且信息不同，程序已停止："]
        for fnsku, row_refs in sorted(duplicate_conflicts.items()):
            lines.append(f"- {fnsku}：{', '.join(row_refs)}")
        raise FatalError("\n".join(lines))

    return {
        "rows": index_rows,
        "warnings": warnings,
        "used_sheets": used_sheet_names,
        "skipped_sheets": skipped_sheet_names,
    }


def urlopen_json(request):
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise FeishuError(f"飞书接口 HTTP {exc.code}：{detail}") from exc
    except urllib.error.URLError as exc:
        raise FeishuError(f"飞书接口连接失败：{exc}") from exc

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise FeishuError(f"飞书接口返回不是 JSON：{raw[:300]}") from exc

    if data.get("code") not in (0, None):
        raise FeishuError(f"飞书接口返回错误：{data}")
    return data


def feishu_json_request(method, path, token=None, payload=None, query=None):
    url = "https://open.feishu.cn" + path
    if query:
        url += "?" + urllib.parse.urlencode(query)

    headers = {
        "User-Agent": "amazon-label-pdf-tool",
        "Accept": "application/json; charset=utf-8",
    }
    body = None
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if payload is not None:
        headers["Content-Type"] = "application/json; charset=utf-8"
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    request = urllib.request.Request(url, data=body, headers=headers, method=method)
    return urlopen_json(request)


def get_secret_value(section, key, default=""):
    try:
        value = st.secrets.get(section, {}).get(key, default)
    except Exception:
        value = default
    return cell_to_text(value)


def load_feishu_config():
    config = {
        "app_id": get_secret_value("feishu", "app_id"),
        "app_secret": get_secret_value("feishu", "app_secret"),
        "spreadsheet_token": get_secret_value("feishu", "spreadsheet_token"),
        "output_folder_token": get_secret_value("feishu", "output_folder_token"),
        "index_range": get_secret_value("feishu", "index_range", FEISHU_INDEX_RANGE_DEFAULT),
    }
    missing = [key for key, value in config.items() if key != "index_range" and not value]
    if missing:
        raise FatalError("Streamlit Secrets 缺少飞书配置：" + "、".join(missing))
    return config


def get_feishu_tenant_access_token(config):
    data = feishu_json_request(
        "POST",
        "/open-apis/auth/v3/tenant_access_token/internal",
        payload={"app_id": config["app_id"], "app_secret": config["app_secret"]},
    )
    token = data.get("tenant_access_token")
    if not token:
        raise FeishuError(f"未获取到 tenant_access_token：{data}")
    return token


def get_nested_value(data, *keys):
    current = data
    for key in keys:
        if not isinstance(current, dict):
            return ""
        current = current.get(key)
    return current


def normalize_feishu_sheet_infos(data):
    payload = data.get("data", {})
    raw_sheets = (
        payload.get("sheets")
        or payload.get("items")
        or payload.get("sheet")
        or payload.get("spreadsheet", {}).get("sheets")
        or []
    )
    if isinstance(raw_sheets, dict):
        raw_sheets = list(raw_sheets.values())

    sheet_infos = []
    for item in raw_sheets:
        if not isinstance(item, dict):
            continue
        title = cell_to_text(
            item.get("title")
            or item.get("name")
            or item.get("sheet_name")
            or item.get("sheetName")
            or get_nested_value(item, "properties", "title")
        )
        sheet_id = cell_to_text(
            item.get("sheet_id")
            or item.get("sheetId")
            or item.get("id")
            or item.get("sheet_token")
            or get_nested_value(item, "properties", "sheet_id")
            or get_nested_value(item, "properties", "sheetId")
        )
        if title or sheet_id:
            sheet_infos.append({"title": title or sheet_id, "sheet_id": sheet_id or title})
    return sheet_infos


def get_feishu_sheet_infos(token, spreadsheet_token):
    errors = []
    for path in (
        f"/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query",
        f"/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/metainfo",
    ):
        try:
            data = feishu_json_request("GET", path, token=token)
            sheet_infos = normalize_feishu_sheet_infos(data)
            if sheet_infos:
                return sheet_infos
        except FeishuError as exc:
            errors.append(str(exc))

    detail = "；".join(errors) if errors else "飞书没有返回工作表列表"
    raise FeishuError(f"读取飞书 Sheet 列表失败：{detail}")


def read_feishu_range_values(token, spreadsheet_token, range_text):
    encoded_range = urllib.parse.quote(range_text, safe="!:$")
    data = feishu_json_request(
        "GET",
        f"/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{encoded_range}",
        token=token,
    )
    value_range = data.get("data", {}).get("valueRange", {})
    return value_range.get("values") or []


def read_feishu_sheet_index(token, config):
    sheet_values_list = []
    for sheet_info in get_feishu_sheet_infos(token, config["spreadsheet_token"]):
        sheet_title = sheet_info["title"]
        if sheet_should_skip(sheet_title):
            sheet_values_list.append((sheet_title, []))
            continue

        range_text = f"{sheet_info['sheet_id']}!{config['index_range']}"
        values = read_feishu_range_values(token, config["spreadsheet_token"], range_text)
        sheet_values_list.append((sheet_title, values))

    return parse_index_sheets(sheet_values_list)


def make_multipart_form(fields, file_field, file_path, display_filename):
    boundary = "----label-feishu-" + uuid.uuid4().hex
    body = bytearray()

    for name, value in fields.items():
        body.extend(f"--{boundary}\r\n".encode())
        body.extend(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode())
        body.extend(str(value).encode("utf-8"))
        body.extend(b"\r\n")

    content_type = mimetypes.guess_type(display_filename)[0] or "application/pdf"
    body.extend(f"--{boundary}\r\n".encode())
    body.extend(
        (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{display_filename}"\r\n'
        ).encode("utf-8")
    )
    body.extend(f"Content-Type: {content_type}\r\n\r\n".encode())
    body.extend(file_path.read_bytes())
    body.extend(b"\r\n")
    body.extend(f"--{boundary}--\r\n".encode())

    return bytes(body), f"multipart/form-data; boundary={boundary}"


def feishu_upload_file(token, config, local_path, display_filename):
    size = local_path.stat().st_size
    fields = {
        "file_name": display_filename,
        "parent_type": "explorer",
        "parent_node": config["output_folder_token"],
        "size": str(size),
    }
    body, content_type = make_multipart_form(fields, "file", local_path, display_filename)
    request = urllib.request.Request(
        "https://open.feishu.cn/open-apis/drive/v1/medias/upload_all",
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": content_type,
            "User-Agent": "amazon-label-pdf-tool",
        },
        method="POST",
    )
    data = urlopen_json(request)
    file_token = data.get("data", {}).get("file_token") or data.get("data", {}).get("token")
    if not file_token:
        raise FeishuError(f"飞书上传成功但未返回 file_token：{data}")
    return file_token


def feishu_list_folder_files(token, config):
    files = []
    page_token = ""
    while True:
        query = {"folder_token": config["output_folder_token"], "page_size": 200}
        if page_token:
            query["page_token"] = page_token

        data = feishu_json_request("GET", "/open-apis/drive/v1/files", token=token, query=query)
        payload = data.get("data", {})
        batch = payload.get("files") or payload.get("items") or []
        files.extend(batch)

        if not payload.get("has_more"):
            break
        page_token = payload.get("next_page_token") or payload.get("page_token") or ""
        if not page_token:
            break
    return files


def feishu_delete_file(token, file_token, file_type="file"):
    if not file_token:
        return
    feishu_json_request(
        "DELETE",
        f"/open-apis/drive/v1/files/{urllib.parse.quote(file_token)}",
        token=token,
        query={"type": file_type},
    )


def build_remote_fnsku_map(remote_files):
    remote_by_fnsku = defaultdict(list)
    pattern = re.compile(FNSKU_REGEX, re.IGNORECASE)
    for item in remote_files:
        name = cell_to_text(item.get("name") or item.get("file_name") or item.get("title"))
        token = cell_to_text(item.get("token") or item.get("file_token"))
        file_type = cell_to_text(item.get("type") or "file")
        if not name or not token:
            continue
        for match in pattern.finditer(name):
            remote_by_fnsku[match.group(0).upper()].append(
                {"name": name, "token": token, "type": file_type}
            )
    return remote_by_fnsku


def extract_fnsku_list(text, pattern):
    found = [match.group(0).upper() for match in pattern.finditer(text)]
    return sorted(set(found))


def natural_sort_key(name):
    parts = re.split(r"(\d+)", name.lower())
    return [int(part) if part.isdigit() else part for part in parts]


def safe_upload_name(filename, fallback_suffix):
    name = Path(filename).name.strip() or f"uploaded{fallback_suffix}"
    name = name.replace("\\", "_").replace("/", "_")
    return name


def unique_path(folder, filename):
    candidate = folder / filename
    if not candidate.exists():
        return candidate

    stem = candidate.stem
    suffix = candidate.suffix
    counter = 1
    while True:
        candidate = folder / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def make_log_row(
    source_name,
    page_number,
    fnsku,
    status,
    reason,
    display_filename,
    local_filename,
    feishu_file_token,
    action,
    label_line1="",
    label_line2="",
):
    return {
        "处理时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "原文件名": source_name,
        "页码": page_number,
        "识别FNSKU": fnsku,
        "匹配状态": status,
        "失败原因": reason,
        "标签第一行": label_line1,
        "标签第二行": label_line2,
        "飞书正式文件名": display_filename,
        "本地备份文件名": local_filename,
        "飞书文件token": feishu_file_token,
        "处理动作": action,
    }


def fail_row(
    source_name,
    category,
    fnsku,
    reason,
    page_number="",
    display_filename="",
    local_filename="",
    label_line1="",
    label_line2="",
):
    return make_log_row(
        source_name,
        page_number,
        fnsku,
        "失败",
        reason,
        display_filename,
        local_filename,
        "",
        f"异常({category}，未上传)",
        label_line1,
        label_line2,
    )


def process_pdf_page(
    pdf_path,
    source_name,
    page_index,
    page_count,
    page_text,
    index_rows,
    fnsku_pattern,
    output_dir,
    batch_fnskus,
    feishu_token,
    feishu_config,
    remote_by_fnsku,
    delete_existing_same_fnsku,
):
    recognized_fnsku = ""
    page_number = page_index + 1

    try:
        text = cell_to_text(page_text)
        if not text:
            raise RuntimeError("无法提取 PDF 文字")

        fnsku_list = extract_fnsku_list(text, fnsku_pattern)
        if not fnsku_list:
            return fail_row(source_name, "未识别FNSKU", "", "未识别 FNSKU", page_number=page_number)

        if len(fnsku_list) > 1:
            recognized_fnsku = "、".join(fnsku_list)
            return fail_row(source_name, "多个FNSKU", recognized_fnsku, "识别到多个不同 FNSKU", page_number=page_number)

        recognized_fnsku = fnsku_list[0]
        if recognized_fnsku in batch_fnskus:
            return fail_row(
                source_name,
                "本次重复FNSKU",
                recognized_fnsku,
                "本次上传已处理过相同 FNSKU，已跳过，避免重复标签",
                page_number=page_number,
            )

        index_row = index_rows.get(recognized_fnsku)
        if not index_row:
            return fail_row(source_name, "索引无匹配", recognized_fnsku, "索引表无匹配", page_number=page_number)

        try:
            label_line1, label_line2 = build_label_lines(index_row)
            display_filename = build_display_filename(index_row, recognized_fnsku)
            local_filename = make_safe_local_filename(display_filename, recognized_fnsku)
        except FileNameError as exc:
            return fail_row(source_name, "信息异常", recognized_fnsku, str(exc), page_number=page_number)

        local_path = unique_path(output_dir, local_filename)
        action = rewrite_label_pdf(
            pdf_path,
            local_path,
            page_index,
            recognized_fnsku,
            label_line1,
            label_line2,
        )

        old_files = remote_by_fnsku.get(recognized_fnsku, [])
        deleted_count = 0
        if delete_existing_same_fnsku:
            for old_file in old_files:
                feishu_delete_file(feishu_token, old_file["token"], old_file.get("type") or "file")
                deleted_count += 1

        feishu_file_token = feishu_upload_file(feishu_token, feishu_config, local_path, display_filename)
        action = f"{action}；上传飞书"
        if deleted_count:
            action += f"；删除飞书旧重复 {deleted_count} 个"

        remote_by_fnsku[recognized_fnsku] = [{"name": display_filename, "token": feishu_file_token, "type": "file"}]
        batch_fnskus.add(recognized_fnsku)

        page_info = f"第 {page_number}/{page_count} 页" if page_count > 1 else str(page_number)
        return make_log_row(
            source_name,
            page_info,
            recognized_fnsku,
            "成功",
            "",
            display_filename,
            local_path.name,
            feishu_file_token,
            action,
            label_line1,
            label_line2,
        )

    except RuntimeError as exc:
        return fail_row(source_name, "无法提取文字", recognized_fnsku, str(exc), page_number=page_number)
    except Exception as exc:
        return fail_row(source_name, "处理失败", recognized_fnsku, f"处理失败：{exc}", page_number=page_number)


def process_pdf_file(
    pdf_path,
    source_name,
    index_rows,
    fnsku_pattern,
    output_dir,
    batch_fnskus,
    feishu_token,
    feishu_config,
    remote_by_fnsku,
    delete_existing_same_fnsku,
):
    try:
        with fitz.open(pdf_path) as document:
            page_count = document.page_count
            if page_count == 0:
                return [fail_row(source_name, "无法提取文字", "", "PDF 没有页面")]
            page_texts = [document[page_index].get_text("text") or "" for page_index in range(page_count)]

        log_rows = []
        for page_index, page_text in enumerate(page_texts):
            log_rows.append(
                process_pdf_page(
                    pdf_path,
                    source_name,
                    page_index,
                    page_count,
                    page_text,
                    index_rows,
                    fnsku_pattern,
                    output_dir,
                    batch_fnskus,
                    feishu_token,
                    feishu_config,
                    remote_by_fnsku,
                    delete_existing_same_fnsku,
                )
            )
        return log_rows

    except Exception as exc:
        return [fail_row(source_name, "处理失败", "", f"处理失败：{exc}")]


def write_log_workbook(log_rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "处理日志"
    worksheet.append(LOG_HEADERS)

    for row in log_rows:
        worksheet.append([row.get(header, "") for header in LOG_HEADERS])

    worksheet.freeze_panes = "A2"
    for column_index, header in enumerate(LOG_HEADERS, start=1):
        values = [header]
        values.extend(cell_to_text(row.get(header, "")) for row in log_rows)
        width = min(max(len(value) for value in values) + 2, 80)
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_zip(output_dir, log_bytes):
    zip_buffer = io.BytesIO()
    log_name = f"处理日志/处理日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("本地备份标签/", b"")
        archive.writestr("处理日志/", b"")
        for pdf_path in sorted(output_dir.glob("*.pdf"), key=lambda path: natural_sort_key(path.name)):
            archive.write(pdf_path, f"本地备份标签/{pdf_path.name}")
        archive.writestr(log_name, log_bytes)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def process_uploads(pdf_files, delete_existing_same_fnsku=True):
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        input_dir = temp_root / "input"
        output_dir = temp_root / "output"
        input_dir.mkdir()
        output_dir.mkdir()

        feishu_config = load_feishu_config()
        feishu_token = get_feishu_tenant_access_token(feishu_config)
        index_result = read_feishu_sheet_index(feishu_token, feishu_config)
        index_rows = index_result["rows"]
        remote_files = feishu_list_folder_files(feishu_token, feishu_config)
        remote_by_fnsku = build_remote_fnsku_map(remote_files)

        saved_pdfs = []
        for uploaded_pdf in pdf_files:
            filename = safe_upload_name(uploaded_pdf.name, ".pdf")
            if not filename.lower().endswith(".pdf"):
                filename = f"{filename}.pdf"
            pdf_path = unique_path(input_dir, filename)
            pdf_path.write_bytes(uploaded_pdf.getvalue())
            saved_pdfs.append((uploaded_pdf.name, pdf_path))

        saved_pdfs.sort(key=lambda item: natural_sort_key(item[0]))

        fnsku_pattern = re.compile(FNSKU_REGEX, re.IGNORECASE)
        log_rows = []
        batch_fnskus = set()

        progress = st.progress(0, text="正在上传到飞书...")
        total = len(saved_pdfs)
        for index, (source_name, pdf_path) in enumerate(saved_pdfs, start=1):
            log_rows.extend(
                process_pdf_file(
                    pdf_path,
                    source_name,
                    index_rows,
                    fnsku_pattern,
                    output_dir,
                    batch_fnskus,
                    feishu_token,
                    feishu_config,
                    remote_by_fnsku,
                    delete_existing_same_fnsku,
                )
            )
            progress.progress(index / total, text=f"正在上传到飞书... {index}/{total}")
        progress.empty()

        log_bytes = write_log_workbook(log_rows)
        zip_bytes = build_zip(output_dir, log_bytes)

    return {
        "warnings": index_result["warnings"],
        "used_sheets": index_result["used_sheets"],
        "skipped_sheets": index_result["skipped_sheets"],
        "remote_fnsku_count": len(remote_by_fnsku),
        "log_rows": log_rows,
        "log_bytes": log_bytes,
        "zip_bytes": zip_bytes,
    }


def render_result(result):
    log_rows = result["log_rows"]
    success_count = sum(1 for row in log_rows if row["匹配状态"] == "成功")
    error_rows = [row for row in log_rows if row["匹配状态"] != "成功"]

    col1, col2, col3 = st.columns(3)
    col1.metric("成功上传", success_count)
    col2.metric("异常标签", len(error_rows))
    col3.metric("总标签", len(log_rows))

    if result["used_sheets"]:
        st.info("已读取索引 Sheet：" + "、".join(result["used_sheets"]))
    if result["skipped_sheets"]:
        st.caption("已跳过非索引 Sheet：" + "、".join(result["skipped_sheets"]))
    for warning in result["warnings"]:
        st.warning(warning)

    if error_rows:
        st.error("有异常标签，异常标签没有上传成功，请查看日志。")
        st.dataframe(
            [
                {
                    "原文件名": row["原文件名"],
                    "页码": row["页码"],
                    "识别FNSKU": row["识别FNSKU"],
                    "失败原因": row["失败原因"],
                }
                for row in error_rows
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success("全部处理并上传成功。")

    reason_counter = Counter(row["失败原因"] for row in error_rows)
    if reason_counter:
        st.caption("异常统计：" + "；".join(f"{reason}：{count}" for reason, count in reason_counter.items()))

    st.download_button(
        "下载处理日志 Excel",
        data=result["log_bytes"],
        file_name=f"处理日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.download_button(
        "下载本次生成 PDF 备份 ZIP",
        data=result["zip_bytes"],
        file_name=f"标签本地备份_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip",
        use_container_width=True,
    )


def main():
    st.set_page_config(page_title="亚马逊标签 PDF 飞书上传工具", layout="wide")
    st.title("亚马逊标签 PDF 飞书上传工具")

    pdf_files = st.file_uploader(
        "上传标签 PDF（可多选，也支持一个文件里有多页标签）",
        type=["pdf"],
        accept_multiple_files=True,
    )

    can_process = bool(pdf_files)
    if st.button("开始处理并上传飞书", type="primary", disabled=not can_process, use_container_width=True):
        try:
            with st.spinner("处理中，请不要关闭页面..."):
                st.session_state["last_result"] = process_uploads(pdf_files, delete_existing_same_fnsku=True)
        except FatalError as exc:
            st.error(f"启动失败：{exc}")
            st.session_state.pop("last_result", None)
        except FeishuError as exc:
            st.error(f"飞书接口失败：{exc}")
            st.session_state.pop("last_result", None)
        except Exception as exc:
            st.error(f"程序异常停止：{exc}")
            st.session_state.pop("last_result", None)

    if "last_result" in st.session_state:
        render_result(st.session_state["last_result"])


if __name__ == "__main__":
    main()
