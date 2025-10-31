# start_agent.ps1
Set-Location "$PSScriptRoot"
& ".\.venv\Scripts\python.exe" -m pip install -q flask python-dotenv
& ".\.venv\Scripts\python.exe" .\local_agent.py
