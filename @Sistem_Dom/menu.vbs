Option Explicit

' Função para exibir o menu e obter a escolha do usuário
Function ExibirMenu()
    Dim escolha
    Do
        escolha = InputBox("Olá" & vbNewLine & _
                           "Qual arquivo você quer rodar?" & vbNewLine & _
                           "1. Info" & vbNewLine & _
                           "2. Lote" & vbNewLine & _
                           "3. Venda" & vbNewLine & _
                           "4. Cadastro" & vbNewLine & _
                           "5. TI" & vbNewLine & _
                           "6. Code" & vbNewLine & _
                           "0. Sair" & vbNewLine & _
                           "Digite o número correspondente ao arquivo que deseja rodar ou 0 para sair:")
        
        If escolha = "" Then
            ' Usuário clicou no botão "Cancelar", sair do loop
            Exit Do
        ElseIf IsNumeric(escolha) Then
            Dim opcao
            opcao = CInt(escolha)
            
            If opcao >= 1 And opcao <= 6 Then
                ExecutarArquivo(opcao)
            ElseIf opcao <> 0 Then
                MsgBox "Opção inválida!"
            End If
        End If
    Loop While escolha <> "0"
End Function

' Função para executar o arquivo correspondente à escolha do usuário
Sub ExecutarArquivo(escolha)
    Dim arquivo
    Select Case escolha
        Case 1
            arquivo = "1-info\info.vbs"
        Case 2
            arquivo = "lote.vbs"
        Case 3
            arquivo = "venda-emp.vbs"
        Case 4
            arquivo = "cadastro-emp.vbs"
        Case 5
            arquivo = "ti-emp.vbs"
        Case 6
            arquivo = "code-emp.vbs"
    End Select
    
    If arquivo <> "" Then
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run arquivo, 1, True
        Set objShell = Nothing
    ElseIf escolha = "0" Then
        MsgBox "Saindo do programa..."
    Else
        MsgBox "Opção inválida!"
    End If
End Sub

' Chamando a função para exibir o menu
ExibirMenu()
