Private Sub Workbook_Open()
    If Date >= "2017-10-1" Then
        Application.DisplayAlerts = False
        MsgBox "��ã���ı���ѵ��ڣ���Ҫ����ʹ��������" & vbCr & "!"'��һ�κ͵ڶ���
        With ThisWorkbook
        .Saved = True
        .ChangeFileAccess xlReadOnly
        Kill .FullName
        .Close
        End With
    End If
End Sub