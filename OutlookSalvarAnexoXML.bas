Attribute VB_Name = "OutlookSalvarAnexoXML"
Sub SalvarAnexosXML()
    Dim objExplorer As Outlook.Explorer
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem As Object
    Dim objAttachment As Outlook.Attachment
    Dim saveFolder As String
    Dim attachmentFileType As String
    
    ' Defina o caminho da pasta onde os anexos .xml serão salvos
    saveFolder = "C:\temp\xml\"
    
    ' Defina o tipo de arquivo para procurar (neste caso, .xml)
    attachmentFileType = ".xml"
    
    ' Obtém o explorador do Outlook ativo
    Set objExplorer = Outlook.Application.ActiveExplorer
    
    ' Verifica se há um item selecionado no explorador
    If Not objExplorer Is Nothing Then
        If objExplorer.Selection.Count > 0 Then
            ' Obtém a pasta do item selecionado
            Set objFolder = objExplorer.Selection(1).Parent
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
            ' Percorre cada anexo no email
            For Each objAttachment In objItem.Attachments
                If Right(objAttachment.FileName, Len(attachmentFileType)) = attachmentFileType Then
                    ' Salva o anexo .xml na pasta especificada
                    objAttachment.SaveAsFile saveFolder & objAttachment.FileName
                End If
            Next objAttachment
        End If
    Next objItem
    
    ' Limpa as variáveis
    Set objAttachment = Nothing
    Set objItem = Nothing
    Set objFolder = Nothing
    Set objExplorer = Nothing
    
    MsgBox "Anexos .xml salvos na pasta " & saveFolder, vbInformation
End Sub
