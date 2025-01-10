```
Sub AtualizarStatusTarefas()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim prazo As Date
    Dim statusAtual As String
    

    Set ws = ThisWorkbook.Sheets("Tarefas")
    
  
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
       
        prazo = ws.Cells(i, 3).Value
        statusAtual = ws.Cells(i, 4).Value
        
        If prazo < Date Then
            ws.Cells(i, 4).Value = "Atrasado"
        ElseIf statusAtual = "Atrasado" Then
            ws.Cells(i, 4).Value = "Pendente"
        End If
    Next i
    
    MsgBox "Status das tarefas atualizado com sucesso!", vbInformation
End Sub

Sub EnviarEmailsTarefas()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim destinatario As String
    Dim descricao As String
    Dim prazo As String
    Dim statusAtual As String
    
  
    Set ws = ThisWorkbook.Sheets("Tarefas")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'Configurando a rota de execução, se der error a aplicação não será quebrada
    'Continuará executando normalmente
    On Error Resume Next
    Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then
        MsgBox "O Outlook não está instalado ou configurado corretamente.", vbExclamation
        Exit Sub
    End If
    
  
    For i = 2 To ultimaLinha
        statusAtual = ws.Cells(i, 4).Value
        
        
        If statusAtual = "Atrasado" Or statusAtual = "Pendente" Then
            destinatario = ws.Cells(i, 6).Value
            descricao = ws.Cells(i, 2).Value
            prazo = ws.Cells(i, 3).Value
            
            'Configurando o corpo da mensagem para envio
            Set MailItem = OutlookApp.CreateItem(0)
            With MailItem
                .To = destinatario
                .Subject = "Lembrete de Tarefa: " & descricao
                .Body = "Olá," & vbCrLf & vbCrLf & _
                        "A tarefa '" & descricao & "' está com o status '" & statusAtual & "' e o prazo é " & prazo & "." & vbCrLf & _
                        "Por favor, tome as medidas necessárias." & vbCrLf & vbCrLf & "Atenciosamente," & vbCrLf & "Equipe de Gerenciamento"
                .Send
            End With
            
            Set MailItem = Nothing
        End If
    Next i
    
    Set OutlookApp = Nothing
    
    MsgBox "E-mails enviados com sucesso!", vbInformation
End Sub

```
