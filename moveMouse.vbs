Set objShell = CreateObject("WScript.Shell")
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objService = objLocator.ConnectServer(".", "root\cimv2")

Do While True
    MoveMouse()
    WScript.Sleep 1000 ' Adjust the delay as necessary
Loop

Sub MoveMouse()
    ' Subroutine to move the mouse cursor
    screenWidth = objService.ExecQuery("SELECT * FROM Win32_DesktopMonitor").ItemIndex(0).ScreenWidth
    screenHeight = objService.ExecQuery("SELECT * FROM Win32_DesktopMonitor").ItemIndex(0).ScreenHeight
    Randomize
    randomX = Int((screenWidth - 1 + 1) * Rnd + 1)
    randomY = Int((screenHeight - 1 + 1) * Rnd + 1)
    objShell.SendKeys "{ESC}" ' Release any keys that might be pressed
    objShell.SendKeys "% " & randomX & " " & randomY ' % indicates Alt key to ensure proper positioning
End Sub
