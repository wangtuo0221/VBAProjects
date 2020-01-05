'需要将Excel文件存成97-03的.xls格式
Private Sub VBAPassword() '
Filename = Application.GetOpenFilename("Excel文件（*.xls & *.xla & *.xlt）,*.xls;*.xla;*.xlt", , "VBA密码")
If Dir(Filename) = "" Then
    MsgBox "未找到相关文件"
    Exit Sub
    Else
    FileCopy Filename, Filename & ".bak" '备份文件。
End If

Dim GetData As String * 5
Open Filename For Binary As #1
Dim CMGs As Long
Dim DPBo As Long

    For i = 1 To LOF(1)
        Get #1, i, GetData
        If GetData = "CMG=""" Then CMGs = i
        If GetData = "[Host" Then DPBo = i - 2: Exit For
    Next

If CMGs = 0 Then
    MsgBox "该工程无密码", 32, "提示"
    Exit Sub
End If

Dim St As String * 2
Dim s20 As String * 1
'取得一个0D0A十六进制字串
Get #1, CMGs - 2, St
'取得一个20十六制字串
Get #1, DPBo + 16, s20
'替换加密部份机码
For i = CMGs To DPBo Step 2
Put #1, i, St
Next
'加入不配对符号
If (DPBo - CMGs) Mod 2 <> 0 Then
    Put #1, DPBo + 1, s20
End If
MsgBox "成功", 32, "提示"
Close #1
End Sub
