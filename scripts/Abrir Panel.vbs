' Abre el panel sin que se vea la consola.
'
' El lanzador .bat corre git antes de abrir, y eso mostraba una ventana de
' consola que aparecia en la barra de tareas con el icono de cmd, al lado del
' panel. Dos entradas para un solo programa, y la de cmd con un icono que no es
' el nuestro.
'
' Minimizarla desde el acceso directo no alcanzaba: igual se registra en la
' barra. Con WindowStyle 0 no se crea ventana en absoluto.
'
' El .bat sigue sirviendo para correrlo a mano y ver que dice el git.

Dim shell, aqui
Set shell = CreateObject("WScript.Shell")
aqui = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

' 0 = sin ventana. False = no esperar a que termine: el panel abre solo.
shell.Run """" & aqui & "Abrir Panel (actualizado).bat""", 0, False
