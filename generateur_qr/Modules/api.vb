Function URL_QRCode_SERIES( _
    ByVal QR_Value As String, _
    Optional ByVal PictureSize As Long = 100, _
    Optional ByVal DisplayText As String = "QRCode >>", _
    Optional ByVal Updateable As Boolean = True) As Variant

    Dim PictureName As String
    Dim oPic As Shape
    Dim oRng As Excel.Range
    Dim vLeft As Variant
    Dim vTop As Variant
    Dim sURL As String

    Const sRootURL As String = "https://chart.googleapis.com/chart?"
    Const sSizeParameter As String = "chs="
    Const sTypeChart As String = "cht=qr"
    Const sDataParameter As String = "chl="
    Const sJoinCHR As String = "&"

    If Not Updateable Then
        URL_QRCode_SERIES = "outdated"
        Exit Function
    End If

    PictureName = "QR-Code_" & DisplayText
    Set oRng = Application.Caller
    On Error Resume Next
    Set oPic = oRng.Parent.Shapes(PictureName)
    On Error GoTo 0

    If oPic Is Nothing Then
        vLeft = oRng.Left + 4
        vTop = oRng.Top
    Else
        vLeft = oPic.Left
        vTop = oPic.Top
        PictureSize = Int(oPic.Width)
        oPic.Delete
    End If

    If Len(QR_Value) = 0 Then
        URL_QRCode_SERIES = CVErr(xlErrValue)
        Exit Function
    End If

    sURL = sRootURL & _
        sSizeParameter & PictureSize & "x" & PictureSize & sJoinCHR & _
        sTypeChart & sJoinCHR & _
        sDataParameter & UTF8_URL_Encode(Replace(QR_Value, " ", "+"))

    Set oPic = oRng.Parent.Shapes.AddPicture(sURL, True, True, vLeft, vTop, PictureSize, PictureSize)
    oPic.Name = PictureName
    URL_QRCode_SERIES = DisplayText
End Function

Private Function URLEncodeByte(val As Integer) As String
    URLEncodeByte = "%" & Right("0" & Hex(val), 2)
End Function

Function UTF8_URL_Encode(ByVal sStr As String) As String
    Dim i As Long
    Dim a As Long
    Dim res As String
    Dim code As String
    
    res = ""
    For i = 1 To Len(sStr)
        a = AscW(Mid(sStr, i, 1))
        If a < 128 Then
            code = Mid(sStr, i, 1)
        ElseIf a < 2048 Then
            code = URLEncodeByte((a \ 64) Or 192) & URLEncodeByte((a And 63) Or 128)
        Else
            code = URLEncodeByte((a \ 144) Or 234) & URLEncodeByte(((a \ 64) And 63) Or 128) & URLEncodeByte((a And 63) Or 128)
        End If
        res = res & code
    Next i
    UTF8_URL_Encode = res
End Function
