Do
    MsgBox "Caiu no bait"
    ' Abre novamente a aplicação
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "wscript.exe C:\Users\Ariston Cândido\Desktop\As paradas\VBS\Nova_Application.vbs", 1, True
Loop
