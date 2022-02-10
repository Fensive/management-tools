set script_location="D:\OneDrive - Fensive Security\OneNote Backup\OneNoteBackup.ps1"
SCHTASKS /CREATE /SC DAILY /TN "FensiveTasks\OneNoteBackup" /TR %script_location% /ST 20:00 /RU %username%