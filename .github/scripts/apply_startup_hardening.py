from pathlib import Path

path = Path("streamlit_app.py")
text = path.read_text(encoding="utf-8")

old_config = '''UPLOAD_WORKERS = max(1, int(os.environ.get("LABEL_TOOL_UPLOAD_WORKERS", "1")))
FEISHU_REQUEST_TIMEOUT = max(5, int(os.environ.get("LABEL_TOOL_FEISHU_TIMEOUT", "45")))'''
new_config = '''def get_int_env(name, default, minimum=1):
    try:
        raw_value = os.environ.get(name)
        value = int(str(raw_value).strip() if raw_value is not None else default)
    except (TypeError, ValueError):
        value = default
    return max(minimum, value)


UPLOAD_WORKERS = get_int_env("LABEL_TOOL_UPLOAD_WORKERS", 1, minimum=1)
FEISHU_REQUEST_TIMEOUT = get_int_env("LABEL_TOOL_FEISHU_TIMEOUT", 45, minimum=5)'''
if old_config not in text and "def get_int_env(" not in text:
    raise SystemExit("config block not found")
text = text.replace(old_config, new_config)

old_reader = '''def write_json_file(path, data):'''
new_reader = '''def safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def write_json_file(path, data):'''
if old_reader not in text and "def safe_float(" not in text:
    raise SystemExit("write_json_file block not found")
text = text.replace(old_reader, new_reader, 1)

text = text.replace(
    'progress_value = float(job.get("progress") or 0)',
    'progress_value = safe_float(job.get("progress"), 0.0)',
)

path.write_text(text, encoding="utf-8")
