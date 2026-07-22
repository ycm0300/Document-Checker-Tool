import asyncio
import hashlib
import json
import re
import shutil
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from config.common_rules import OUTPUT_DIR
from web_service import check_word_files, workbook_to_sheets


WEB_JOBS_DIR = OUTPUT_DIR / "web_jobs"
WEB_DOCUMENTS_DIR = OUTPUT_DIR / "web_documents"
MAX_FILE_SIZE = 50 * 1024 * 1024
JOB_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")
CHECK_PROGRESS = {}
PERSISTENCE_LOCK = threading.Lock()

app = FastAPI(title="Document Compliance Checker API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


def get_job_dir(job_id):
    if not JOB_ID_PATTERN.fullmatch(job_id):
        raise HTTPException(status_code=404, detail="任务不存在")
    return WEB_JOBS_DIR / job_id


def unique_upload_path(upload_dir, original_name):
    safe_name = Path(original_name or "document.docx").name
    path = upload_dir / safe_name
    counter = 2
    while path.exists():
        path = upload_dir / f"{Path(safe_name).stem}_{counter}{Path(safe_name).suffix}"
        counter += 1
    return path


def build_file_signature(files):
    """按文件名组合识别网页任务，忽略选择顺序和大小写。"""
    return sorted(Path(file.filename or "document.docx").name.casefold() for file in files)


def find_existing_jobs(file_signature):
    if not WEB_JOBS_DIR.exists():
        return []

    matches = []
    for job_dir in WEB_JOBS_DIR.iterdir():
        if not job_dir.is_dir() or not JOB_ID_PATTERN.fullmatch(job_dir.name):
            continue
        metadata_file = job_dir / "metadata.json"
        stored_signature = None
        if metadata_file.is_file():
            try:
                metadata = json.loads(metadata_file.read_text(encoding="utf-8"))
                stored_signature = metadata.get("fileSignature")
            except (OSError, json.JSONDecodeError):
                pass
        else:
            upload_dir = job_dir / "uploads"
            if upload_dir.is_dir():
                stored_signature = sorted(file.name.casefold() for file in upload_dir.glob("*.docx"))

        if stored_signature == file_signature and (job_dir / "总检查结果.xlsx").is_file():
            matches.append(job_dir)

    return sorted(matches, key=lambda path: path.stat().st_mtime, reverse=True)


def replace_existing_job(new_job_dir, existing_job_id):
    """成功生成新结果后，以可恢复方式替换同文件组合的旧任务。"""
    existing_dir = WEB_JOBS_DIR / existing_job_id
    backup_dir = WEB_JOBS_DIR / f".{existing_job_id}.backup-{uuid.uuid4().hex}"
    existing_dir.rename(backup_dir)
    try:
        new_job_dir.rename(existing_dir)
    except Exception:
        backup_dir.rename(existing_dir)
        raise
    shutil.rmtree(backup_dir, ignore_errors=True)
    return existing_dir


def document_storage_key(file_name):
    normalized_name = Path(file_name).name.casefold().encode("utf-8")
    return hashlib.sha256(normalized_name).hexdigest()[:24]


def replace_document_record(source_file, issues):
    """只替换当前已检查文档的独立归档，不触碰其他文档。"""
    source_file = Path(source_file)
    WEB_DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)
    storage_key = document_storage_key(source_file.name)
    target_dir = WEB_DOCUMENTS_DIR / storage_key
    staging_dir = WEB_DOCUMENTS_DIR / f".{storage_key}.staging-{uuid.uuid4().hex}"
    backup_dir = WEB_DOCUMENTS_DIR / f".{storage_key}.backup-{uuid.uuid4().hex}"
    staging_dir.mkdir(parents=True)

    try:
        shutil.copy2(source_file, staging_dir / source_file.name)
        (staging_dir / "issues.json").write_text(
            json.dumps(issues, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (staging_dir / "metadata.json").write_text(
            json.dumps({
                "fileName": source_file.name,
                "updatedAt": datetime.now(timezone.utc).isoformat(),
            }, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        if target_dir.exists():
            target_dir.rename(backup_dir)
        staging_dir.rename(target_dir)
        if backup_dir.exists():
            shutil.rmtree(backup_dir, ignore_errors=True)
    except Exception:
        shutil.rmtree(staging_dir, ignore_errors=True)
        if backup_dir.exists() and not target_dir.exists():
            backup_dir.rename(target_dir)
        raise


@app.get("/api/health")
def health_check():
    return {"status": "ok"}


def update_check_progress(progress_id, percent, message, current_file=None):
    if not progress_id:
        return
    CHECK_PROGRESS[progress_id] = {
        "percent": max(0, min(100, round(percent))),
        "message": message,
        "currentFile": current_file,
    }


@app.get("/api/check-progress/{progress_id}")
def get_check_progress(progress_id: str):
    if not JOB_ID_PATTERN.fullmatch(progress_id) or progress_id not in CHECK_PROGRESS:
        raise HTTPException(status_code=404, detail="检查进度不存在")
    return CHECK_PROGRESS[progress_id]


@app.post("/api/check")
async def check_documents(
    files: list[UploadFile] = File(...),
    x_progress_id: str | None = Header(default=None),
):
    progress_id = x_progress_id if x_progress_id and JOB_ID_PATTERN.fullmatch(x_progress_id) else None
    if progress_id:
        if len(CHECK_PROGRESS) >= 200:
            CHECK_PROGRESS.pop(next(iter(CHECK_PROGRESS)))
        update_check_progress(progress_id, 1, "准备上传文档")
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个 Word 文档")

    invalid_files = [file.filename for file in files if Path(file.filename or "").suffix.lower() != ".docx"]
    if invalid_files:
        raise HTTPException(status_code=400, detail=f"仅支持 .docx 文件：{', '.join(invalid_files)}")

    file_signature = build_file_signature(files)
    existing_jobs = find_existing_jobs(file_signature)
    existing_job_id = existing_jobs[0].name if existing_jobs else None
    temporary_job_id = uuid.uuid4().hex
    job_id = existing_job_id or temporary_job_id
    job_dir = WEB_JOBS_DIR / temporary_job_id
    upload_dir = job_dir / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=False)
    saved_files = []

    try:
        for upload_index, upload in enumerate(files):
            update_check_progress(
                progress_id,
                5 + (upload_index / len(files)) * 15,
                f"正在上传 {upload.filename}",
                upload.filename,
            )
            target = unique_upload_path(upload_dir, upload.filename)
            size = 0
            with target.open("wb") as destination:
                while chunk := await upload.read(1024 * 1024):
                    size += len(chunk)
                    if size > MAX_FILE_SIZE:
                        raise HTTPException(status_code=413, detail=f"文件不能超过 50 MB：{upload.filename}")
                    destination.write(chunk)
            saved_files.append(target)
            await upload.close()

        def report_file_progress(completed, total, current_file, stage):
            percent = 20 + (completed / max(total, 1)) * 65
            message = f"{stage}：{current_file}" if current_file else stage
            update_check_progress(progress_id, percent, message, current_file)

        excel_file, all_file_issues = await asyncio.to_thread(
            check_word_files,
            saved_files,
            job_dir,
            report_file_progress,
        )
        update_check_progress(progress_id, 90, "正在整理检查结果")
        sheet_data = workbook_to_sheets(excel_file)

        # 多个检查请求可能同时更新同名文档或同一文件组合。
        # Windows 不允许将目录重命名到已被另一请求创建的目标目录，
        # 因此将归档和任务替换作为一个互斥的持久化阶段。
        with PERSISTENCE_LOCK:
            for saved_file in saved_files:
                replace_document_record(saved_file, all_file_issues[saved_file.name])
            metadata = {
                "jobId": job_id,
                "fileSignature": file_signature,
                "documentNames": [Path(file.filename or "document.docx").name for file in files],
                "updatedAt": datetime.now(timezone.utc).isoformat(),
            }
            (job_dir / "metadata.json").write_text(
                json.dumps(metadata, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            if existing_job_id:
                job_dir = replace_existing_job(job_dir, existing_job_id)
                excel_file = job_dir / excel_file.name
                for duplicate_job in existing_jobs[1:]:
                    shutil.rmtree(duplicate_job, ignore_errors=True)

        update_check_progress(progress_id, 100, "检查完成")
        return {
            "jobId": job_id,
            "fileName": excel_file.name,
            "documentCount": len(saved_files),
            "sheets": sheet_data,
            "excelDownloadUrl": f"/api/jobs/{job_id}/excel",
            "replacedExistingTask": existing_job_id is not None,
        }
    except HTTPException:
        update_check_progress(progress_id, 100, "检查失败")
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as error:
        update_check_progress(progress_id, 100, "检查失败")
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"检查失败：{error}") from error


@app.get("/api/jobs/{job_id}/excel")
def download_excel(job_id: str):
    excel_file = get_job_dir(job_id) / "总检查结果.xlsx"
    if not excel_file.is_file():
        raise HTTPException(status_code=404, detail="检查结果不存在或已被清理")
    return FileResponse(
        excel_file,
        filename="document-compliance-results.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/jobs/{job_id}")
def get_job_result(job_id: str):
    excel_file = get_job_dir(job_id) / "总检查结果.xlsx"
    if not excel_file.is_file():
        raise HTTPException(status_code=404, detail="检查结果不存在或已被清理")
    return {
        "jobId": job_id,
        "fileName": excel_file.name,
        "sheets": workbook_to_sheets(excel_file),
        "excelDownloadUrl": f"/api/jobs/{job_id}/excel",
    }
