from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
APP_DIR=ROOT/"app";CONFIG_DIR=ROOT/"config";DATA_DIR=ROOT/"data";IMPORT_DIR=DATA_DIR/"imports";BACKUP_DIR=DATA_DIR/"backups"
LOG_DIR=ROOT/"logs";ERROR_DIR=LOG_DIR/"errors";RUNTIME_DIR=ROOT/"runtime";WEB_DIR=APP_DIR/"web";STATIC_DIR=WEB_DIR/"static"
DB_PATH=DATA_DIR/"app.db";AUDIT_JSONL=LOG_DIR/"audit.jsonl";APP_LOG=LOG_DIR/"app.log";BROWSER_LOG=LOG_DIR/"browser-actions.jsonl"
for p in (DATA_DIR,IMPORT_DIR,BACKUP_DIR,LOG_DIR,ERROR_DIR,RUNTIME_DIR,RUNTIME_DIR/"browser_profile"):p.mkdir(parents=True,exist_ok=True)
