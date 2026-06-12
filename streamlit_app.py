import io
import re
import tempfile
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from string import Formatter

import fitz
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


FNSKU_REGEX = r"\bX[A-Z0-9]{9}\b"
FILENAME_TEMPLATE = "{SKU}-{FNSKU}.pdf"
LABEL_TEXT_LINE2_TEMPLATE = "{款式}"

PDF_TEXT_COVER_TOP_RATIO = 0.535
PDF_FNSKU_BASELINE_RATIO = 0.660
PDF_LINE1_BASELINE_RATIO = 0.783
PDF_LINE2_BASELINE_RATIO = 0.895
PDF_SIDE_MARGIN = 4

# 使用 PyMuPDF 内置中文字体，不嵌入系统字体，文件小、速度快。
CJK_FONT_CANDIDATES = []
CJK_BOLD_STROKE_WIDTH = 0.05

LOG_HEADERS = [
    "处理时间",
    "原文件名",
    "页码",
    "识别FNSKU",
    "匹配状态",
    "失败原因",
    "标签第一行",
    "标签第二行",
    "输出文件名",
    "输出路径",
    "处理动作",
]

REQUIRED_INDEX_HEADERS = ["FNSKU", "SKU", "款式", "品牌", "型号", "产品类型"]
REQUIRED_INFO_FIELDS = ["SKU", "款式", "品牌", "型号", "产品类型"]
OPTIONAL_INFO_FIELDS = ["颜色"]
ILLEGAL_FILENAME_CHARS = r'\:*?"<>|'


class FatalError(Exception):
    """启动阶段错误：继续处理会导致结果不可靠。"""


class FileNameError(Exception):
    """文件名无法安全生成。"""


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


def sanitize_filename(filename):
    filename = cell_to_text(filename)
    filename = re.sub(r"\s+", "-", filename)

    for char in ILLEGAL_FILENAME_CHARS:
        filename = filename.replace(char, "-")

    filename = filename.replace("/", "")
    filename = re.sub(r"-+", "-", filename)
    filename = filename.strip(" .-")
    if not filename:
        raise FileNameError("文件名为空")

    filename = ensure_pdf_suffix(filename)
    if filename.lower() == ".pdf":
        raise FileNameError("文件名为空")

    return filename


def get_template_fields(template):
    fields = []
    for _, field_name, _, _ in Formatter().parse(template):
        if not field_name:
            continue
        simple_name = field_name.split(".", 1)[0].split("[", 1)[0]
        if simple_name and simple_name not in fields:
            fields.append(simple_name)
    return fields


def build_output_filename(row, fnsku):
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
        filename = sanitize_filename(FILENAME_TEMPLATE.format(**values))
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
    temp_path = target_path.parent / f".{target_path.stem}.tmp.pdf"
    if temp_path.exists():
        temp_path.unlink()

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

        temp_path.replace(target_path)
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise

    return "生成修改后PDF"


def read_index_file(index_path):
    workbook = load_workbook(index_path, data_only=True)
    worksheet = workbook.active

    header_values = [
        cell_to_text(value) for value in next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True))
    ]
    header_to_index = {header: idx for idx, header in enumerate(header_values) if header}

    missing_headers = [header for header in REQUIRED_INDEX_HEADERS if header not in header_to_index]
    if missing_headers:
        raise FatalError("Excel 索引表缺少表头：" + "、".join(missing_headers))

    missing_fnsku_rows = []
    invalid_fnsku_rows = []
    skipped_same_duplicate_rows = []
    index_rows = {}
    duplicate_conflicts = defaultdict(list)
    duplicate_check_fields = REQUIRED_INFO_FIELDS + OPTIONAL_INFO_FIELDS

    for row_number, values in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
        row = {}
        has_any_value = False

        for header, idx in header_to_index.items():
            value = values[idx] if idx < len(values) else None
            text = cell_to_text(value)
            row[header] = text
            if text:
                has_any_value = True

        if not has_any_value:
            continue

        fnsku = normalize_fnsku(row.get("FNSKU"))
        if not fnsku:
            missing_fnsku_rows.append(row_number)
            continue

        if not is_valid_fnsku(fnsku):
            invalid_fnsku_rows.append((row_number, fnsku))
            continue

        row["FNSKU"] = fnsku
        row["_Excel行号"] = row_number

        existing_row = index_rows.get(fnsku)
        if existing_row:
            is_same_content = all(
                cell_to_text(existing_row.get(field)) == cell_to_text(row.get(field))
                for field in duplicate_check_fields
            )
            if is_same_content:
                skipped_same_duplicate_rows.append((row_number, fnsku, existing_row["_Excel行号"]))
                continue

            if not duplicate_conflicts[fnsku]:
                duplicate_conflicts[fnsku].append(existing_row["_Excel行号"])
            duplicate_conflicts[fnsku].append(row_number)
            continue

        index_rows[fnsku] = row

    if missing_fnsku_rows:
        raise FatalError(
            "Excel 中以下行缺少必填 FNSKU："
            + "、".join(str(row_number) for row_number in missing_fnsku_rows)
        )

    if duplicate_conflicts:
        lines = ["Excel 中 FNSKU 重复且信息不同，程序已停止："]
        for fnsku, row_numbers in sorted(duplicate_conflicts.items()):
            lines.append(f"- {fnsku}：第 {', '.join(str(row) for row in row_numbers)} 行")
        raise FatalError("\n".join(lines))

    warnings = []
    if invalid_fnsku_rows:
        sample_text = "；".join(
            f"第 {row_number} 行 {fnsku}" for row_number, fnsku in invalid_fnsku_rows[:10]
        )
        if len(invalid_fnsku_rows) > 10:
            sample_text += "；..."
        warnings.append(f"已忽略 {len(invalid_fnsku_rows)} 行无效 FNSKU（{sample_text}）")

    if skipped_same_duplicate_rows:
        sample_text = "；".join(
            f"{fnsku} 第 {row_number} 行（与第 {first_row_number} 行一致）"
            for row_number, fnsku, first_row_number in skipped_same_duplicate_rows[:10]
        )
        if len(skipped_same_duplicate_rows) > 10:
            sample_text += "；..."
        warnings.append(f"已忽略 {len(skipped_same_duplicate_rows)} 行完全相同的重复 FNSKU（{sample_text}）")

    return index_rows, warnings


def extract_fnsku_list(text, pattern):
    found = [match.group(0).upper() for match in pattern.finditer(text)]
    return sorted(set(found))


def make_log_row(
    source_name,
    page_number,
    fnsku,
    status,
    reason,
    output_filename,
    output_path,
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
        "输出文件名": output_filename,
        "输出路径": output_path,
        "处理动作": action,
    }


def fail_row(
    source_name,
    category,
    fnsku,
    reason,
    page_number="",
    output_filename="",
    label_line1="",
    label_line2="",
):
    return make_log_row(
        source_name,
        page_number,
        fnsku,
        "失败",
        reason,
        output_filename,
        "",
        f"异常({category}，未生成PDF)",
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
    generated_fnskus,
    generated_filenames,
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
            return fail_row(
                source_name,
                "多个FNSKU",
                recognized_fnsku,
                "识别到多个不同 FNSKU",
                page_number=page_number,
            )

        recognized_fnsku = fnsku_list[0]
        index_row = index_rows.get(recognized_fnsku)
        if not index_row:
            return fail_row(source_name, "Excel无匹配", recognized_fnsku, "Excel 无匹配", page_number=page_number)

        try:
            label_line1, label_line2 = build_label_lines(index_row)
            output_filename = build_output_filename(index_row, recognized_fnsku)
        except FileNameError as exc:
            return fail_row(source_name, "信息异常", recognized_fnsku, str(exc), page_number=page_number)

        if recognized_fnsku in generated_fnskus:
            return fail_row(
                source_name,
                "重复FNSKU",
                recognized_fnsku,
                "本次上传中已生成过相同 FNSKU，已跳过，避免重复标签",
                page_number=page_number,
                output_filename=output_filename,
                label_line1=label_line1,
                label_line2=label_line2,
            )

        if output_filename.lower() in generated_filenames:
            return fail_row(
                source_name,
                "文件名冲突",
                recognized_fnsku,
                "本次上传中输出文件名冲突，已跳过",
                page_number=page_number,
                output_filename=output_filename,
                label_line1=label_line1,
                label_line2=label_line2,
            )

        output_path = output_dir / output_filename
        action = rewrite_label_pdf(
            pdf_path,
            output_path,
            page_index,
            recognized_fnsku,
            label_line1,
            label_line2,
        )

        generated_fnskus.add(recognized_fnsku)
        generated_filenames.add(output_filename.lower())

        page_info = f"第 {page_number}/{page_count} 页" if page_count > 1 else str(page_number)
        return make_log_row(
            source_name,
            page_info,
            recognized_fnsku,
            "成功",
            "",
            output_filename,
            f"已重命名标签库/{output_filename}",
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
    generated_fnskus,
    generated_filenames,
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
                    generated_fnskus,
                    generated_filenames,
                )
            )
        return log_rows

    except Exception as exc:
        return [fail_row(source_name, "处理失败", "", f"处理失败：{exc}")]


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


def make_template_workbook():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "标签索引表"
    worksheet.append(["FNSKU", "SKU", "款式", "品牌", "型号", "产品类型", "颜色"])
    worksheet.append([
        "X001B4BRHJ",
        "AAABBBCCC",
        "CASEME-背面卡包",
        "ELEPIK",
        "iPhone 17",
        "Case",
        "Fashion Purple",
    ])

    widths = [16, 28, 22, 14, 20, 16, 20]
    for column_index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column_index)].width = width

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_zip(output_dir, log_bytes):
    zip_buffer = io.BytesIO()
    log_name = f"处理日志/处理日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("已重命名标签库/", b"")
        archive.writestr("异常待核对/", b"")
        archive.writestr("处理日志/", b"")
        for pdf_path in sorted(output_dir.glob("*.pdf"), key=lambda path: natural_sort_key(path.name)):
            archive.write(pdf_path, f"已重命名标签库/{pdf_path.name}")
        archive.writestr(log_name, log_bytes)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def process_uploads(index_file, pdf_files):
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        input_dir = temp_root / "input"
        output_dir = temp_root / "output"
        input_dir.mkdir()
        output_dir.mkdir()

        index_path = temp_root / "标签索引表.xlsx"
        index_path.write_bytes(index_file.getvalue())

        index_rows, warnings = read_index_file(index_path)

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
        generated_fnskus = set()
        generated_filenames = set()

        progress = st.progress(0, text="正在处理 PDF...")
        total = len(saved_pdfs)
        for index, (source_name, pdf_path) in enumerate(saved_pdfs, start=1):
            log_rows.extend(
                process_pdf_file(
                    pdf_path,
                    source_name,
                    index_rows,
                    fnsku_pattern,
                    output_dir,
                    generated_fnskus,
                    generated_filenames,
                )
            )
            progress.progress(index / total, text=f"正在处理 PDF... {index}/{total}")
        progress.empty()

        log_bytes = write_log_workbook(log_rows)
        zip_bytes = build_zip(output_dir, log_bytes)

    return {
        "warnings": warnings,
        "log_rows": log_rows,
        "log_bytes": log_bytes,
        "zip_bytes": zip_bytes,
    }


def render_result(result):
    log_rows = result["log_rows"]
    success_count = sum(1 for row in log_rows if row["匹配状态"] == "成功")
    error_rows = [row for row in log_rows if row["匹配状态"] != "成功"]

    col1, col2, col3 = st.columns(3)
    col1.metric("成功标签", success_count)
    col2.metric("异常标签", len(error_rows))
    col3.metric("总标签", len(log_rows))

    for warning in result["warnings"]:
        st.warning(warning)

    if error_rows:
        st.error("有异常标签，异常文件未放入成功文件夹，请查看日志。")
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
        st.success("全部处理成功。")

    if log_rows:
        reason_counter = Counter(row["失败原因"] for row in error_rows)
        if reason_counter:
            st.caption("异常统计：" + "；".join(f"{reason}：{count}" for reason, count in reason_counter.items()))

    st.download_button(
        "下载处理结果 ZIP",
        data=result["zip_bytes"],
        file_name=f"标签处理结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip",
        use_container_width=True,
    )

    st.download_button(
        "只下载处理日志 Excel",
        data=result["log_bytes"],
        file_name=f"处理日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def main():
    st.set_page_config(page_title="亚马逊标签PDF自动整理工具", layout="wide")

    st.title("亚马逊标签PDF自动整理工具")

    with st.sidebar:
        st.subheader("索引表模板")
        st.download_button(
            "下载模板",
            data=make_template_workbook(),
            file_name="标签索引表模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    index_file = st.file_uploader("上传标签索引表（.xlsx）", type=["xlsx"])
    pdf_files = st.file_uploader("上传标签 PDF（可多选，支持单个多页 PDF）", type=["pdf"], accept_multiple_files=True)

    can_process = index_file is not None and bool(pdf_files)

    if st.button("开始处理", type="primary", disabled=not can_process, use_container_width=True):
        try:
            with st.spinner("处理中，请不要关闭页面..."):
                st.session_state["last_result"] = process_uploads(index_file, pdf_files)
        except FatalError as exc:
            st.error(f"启动失败：{exc}")
            st.session_state.pop("last_result", None)
        except Exception as exc:
            st.error(f"程序异常停止：{exc}")
            st.session_state.pop("last_result", None)

    if "last_result" in st.session_state:
        render_result(st.session_state["last_result"])


if __name__ == "__main__":
    main()
