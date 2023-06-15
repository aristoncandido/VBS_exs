Dim title
title = "Tela de Login"

Dim email
Dim pass

Set objShell = CreateObject("WScript.Shell")

' Executa o Bloco de Notas
objShell.Run "notepad.exe", 1, True

' Aguarda o Bloco de Notas abrir
WScript.Sleep 1000

' Envia as teclas para o Bloco de Notas
objShell.SendKeys email & "{ENTER}"

' Aguarda um pouco para a caixa de diálogo ser exibida
WScript.Sleep 500

' Captura o conteúdo da caixa de diálogo de entrada
objShell.SendKeys "^a"
objShell.SendKeys "^c"
email = objShell.Clipboard.GetText()

' Limpa o conteúdo da caixa de diálogo do Bloco de Notas
objShell.SendKeys "^a"
objShell.SendKeys "{DEL}"

' Exibe o valor capturado
MsgBox "Email: " & email, vbInformation, title
