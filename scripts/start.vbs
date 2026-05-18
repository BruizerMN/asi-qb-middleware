' Launches the ASI QB Middleware Python app with no console window.
' Called by the Task Scheduler logon task.
Dim shell, repoPath
repoPath = "C:\Services\asi-qb-middleware"
Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = repoPath
shell.Run "py app.py", 0, False
