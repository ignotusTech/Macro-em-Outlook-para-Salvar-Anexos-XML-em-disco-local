Attribute VB_Name = "OutlookSalvarAnexoXML"
Sub SalvarAnexosXML()
    Dim objExplorer As Outlook.Explorer
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem As Object
    Dim objAttachment As Outlook.Attachment
    Dim saveFolder As String
    Dim attachmentFileType As String
    Dim contTotal As Long ' Variável para contar o número total de emails verificados
    Dim contXML As Long ' Variável para contar o número de anexos .xml salvos
    
    ' Inicializa as variáveis de contagem
    contTotal = 0
    contXML = 0
    
    ' Defina o caminho da pasta onde os anexos .xml serão salvos
    saveFolder = "c:\temp\"
    
    ' Defina o tipo de arquivo para procurar (neste caso, .xml)
    attachmentFileType = ".xml"
    
    ' Obtém o explorador do Outlook ativo
    Set objExplorer = Outlook.Application.ActiveExplorer
    
    ' Verifica se há um item selecionado no explorador
    If Not objExplorer Is Nothing Then
        If objExplorer.Selection.Count > 0 Then
            ' Obtém a pasta do item selecionado
            Set objFolder = objExplorer.CurrentFolder
        End If
    End If
    
    ' Se não houver uma pasta selecionada, saia da macro
    If objFolder Is Nothing Then
        MsgBox "Nenhuma pasta selecionada.", vbExclamation
        Exit Sub
    End If
    
    ' Percorre cada email na pasta selecionada
    For Each objItem In objFolder.Items
        If objItem.Class = olMail Then
            ' Incrementa o contador de emails verificados
            contTotal = contTotal + 1
            ' Percorre cada anexo no email
            For Each objAttachment In objItem.Attachments
                Dim fileExt As String
                fileExt = Right(objAttachment.FileName, Len(attachmentFileType))
                ' Verifica se a extensão do arquivo corresponde (insensível a maiúsculas e minúsculas)
                If StrComp(fileExt, attachmentFileType, vbTextCompare) = 0 Then
                    ' Salva o anexo .xml na pasta especificada
                    On Error Resume Next
                    objAttachment.SaveAsFile saveFolder & objAttachment.FileName
                    On Error GoTo 0 ' Desativa o tratamento de erros
                    ' Verifica se houve um erro ao salvar o anexo
                    ' Verifique se ocorreu um erro ao salvar o anexo
                    If Err.Number <> 0 Then
                        ' Registre o erro em algum lugar, como no Immediate Window (janela imediata)
                        Debug.Print "Erro ao salvar o anexo: " & Err.Description
                        Err.Clear ' Limpe o objeto de erro
                    Else
                        ' Incrementa o contador de anexos .xml salvos
                        contXML = contXML + 1
                    End If
                End If
            Next objAttachment
        End If
    Next objItem
    

    ' Limpa as variáveis
    Set objAttachment = Nothing
    Set objItem = Nothing
    Set objFolder = Nothing
    Set objExplorer = Nothing
    
    ' Exibe a contagem no final
    MsgBox "Foram verificados " & contTotal & " emails e " & contXML & " anexos .xml foram salvos na pasta " & saveFolder, vbInformation
End Sub
