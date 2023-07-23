Option Explicit

' Cria a interface da caixa de mensagem personalizada
Dim objShell, intResponse
Set objShell = CreateObject("WScript.Shell")

intResponse = MsgBox("Aqui está uma caixa de diálogo mais bonita." & vbCrLf & _
                     "Clique em 'OK' para fechar o script.", vbOKOnly + vbInformation, "Caixa de Diálogo")

' Fecha o script
WScript.Quit
