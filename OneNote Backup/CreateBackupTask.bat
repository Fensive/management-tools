set script_location="D:\OneDrive - Fensive Security\management-tools\OneNote Backup\OneNoteBackup.ps1"
SCHTASKS /CREATE /SC DAILY /TN "FensiveTasks\OneNoteBackup" /TR "powershell -executionpolicy bypass -windowstyle Hidden -File ""%script_location%""" /ST 20:00 /RU %username%
