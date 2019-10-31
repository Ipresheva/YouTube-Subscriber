set WshShell = WScript.CreateObject("WScript.Shell")
i = 0
do
WshShell.SendKeys("{TAB}")
i = i + 1
loop while i < 10
WshShell.SendKeys("{ENTER}")
