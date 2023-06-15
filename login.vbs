Dim title
title = "Tela de Login"

Dim email
Dim pass

Set objShell = CreateObject("WScript.Shell")
email = objShell.InputBox("Digite seu Email:", title, "");
