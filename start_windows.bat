@echo off
REM รันบน Windows โดยไม่ใช้ Docker (ต้องติดตั้ง Python 3.10+ ก่อน)
cd /d %~dp0
if not exist .venv ( python -m venv .venv )
call .venv\Scripts\activate
pip install -r requirements.txt
REM ตั้งรหัสเข้าเว็บ (ลบ REM ข้างล่างแล้วเปลี่ยนเลข) - ควรตั้งเมื่อเปิดเป็นสาธารณะ
REM set ACCESS_CODE=1234
REM set ADMIN_PIN=9999
cd app
python -m uvicorn main:app --host 0.0.0.0 --port 8000
