from pathlib import Path

path = Path("streamlit_app.py")
text = path.read_text(encoding="utf-8")

text = text.replace(
    'UPLOAD_WORKERS = max(1, int(os.environ.get("LABEL_TOOL_UPLOAD_WORKERS", "5")))',
    'UPLOAD_WORKERS = max(1, int(os.environ.get("LABEL_TOOL_UPLOAD_WORKERS", "1")))',
)

old_cancel = '''def request_job_cancel(job_dir):
    job_cancel_path(job_dir).write_text(datetime.now().isoformat(), encoding="utf-8")
    return update_job_metadata(
        job_dir,
        cancel_requested=True,
        status="正在停止",
        message="已请求停止，正在结束任务，不会继续上传后续文件。",
    )


def prune_job_history():'''
new_cancel = '''def request_job_cancel(job_dir):
    job_cancel_path(job_dir).write_text(datetime.now().isoformat(), encoding="utf-8")
    return update_job_metadata(
        job_dir,
        cancel_requested=True,
        status="正在停止",
        message="已请求停止，正在结束任务，不会继续上传后续文件。",
    )


def clear_finished_job_history():
    if not JOB_DIR.exists():
        return 0

    removed_count = 0
    active_statuses = {"排队中", "处理中", "正在停止"}
    for job_dir in [path for path in JOB_DIR.iterdir() if path.is_dir()]:
        metadata = read_json_file(job_metadata_path(job_dir), {}) or {}
        if metadata.get("status") in active_statuses:
            continue
        shutil.rmtree(job_dir, ignore_errors=True)
        removed_count += 1
    return removed_count


def prune_job_history():'''
if old_cancel not in text and "def clear_finished_job_history():" not in text:
    raise SystemExit("cancel block not found")
text = text.replace(old_cancel, new_cancel)

old_header = '''    st.divider()
    cols = st.columns([1, 4])
    with cols[0]:
        st.button("刷新任务状态", use_container_width=True)
    with cols[1]:
        st.subheader("后台任务")
        st.caption("页面不会自动刷新。需要看最新进度时，点左侧刷新按钮即可。")'''
new_header = '''    st.divider()
    cols = st.columns([1, 1, 4])
    with cols[0]:
        st.button("刷新任务状态", use_container_width=True)
    with cols[1]:
        if st.button("清空已结束记录", use_container_width=True):
            removed_count = clear_finished_job_history()
            st.success(f"已清空 {removed_count} 条已结束任务记录。")
            st.rerun()
    with cols[2]:
        st.subheader("后台任务")
        st.caption("页面不会自动刷新。需要看最新进度时，点左侧刷新按钮即可。")'''
if old_header not in text and "清空已结束记录" not in text:
    raise SystemExit("job header block not found")
text = text.replace(old_header, new_header)

text = text.replace(
    'st.warning("已发出停止请求。当前正在上传的几个文件会先收尾，后面的不会继续上传。")',
    'st.warning("已发出停止请求。任务会尽快结束，不会继续上传后续文件。")',
)
text = text.replace(
    'st.warning("正在停止：等待当前上传请求收尾。")',
    'st.warning("正在停止：任务正在结束，不会继续上传后续文件。")',
)

path.write_text(text, encoding="utf-8")
