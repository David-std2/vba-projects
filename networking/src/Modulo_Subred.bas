Sub Subred()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim clrR As Long
    clrR = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If clrR >= 4 Then
        With ws.Range("A4:H" & clrR + 500)
            .UnMerge
            .Clear
        End With
    End If
    
    Dim ip As String
    Dim mBase As Integer
    Dim mObj As Integer
    
    ip = ws.Range("A1").Value
    If ip = "" Then Exit Sub
    
    mBase = ws.Range("B1").Value
    
    If Trim(ws.Range("C1").Value) = "" Then
        mObj = mBase
    Else
        mObj = ws.Range("C1").Value
    End If
    
    If mObj < mBase Then
        MsgBox "El prefijo objetivo debe ser mayor o igual al prefijo base.", vbCritical
        Exit Sub
    End If
    
    Dim ipD As Double
    ipD = IpToDbl(ip)
    
    Dim bS As Double
    bS = 2 ^ (32 - mBase)
    ipD = Int(ipD / bS) * bS
    
    Dim bT As Double
    bT = 2 ^ (32 - mObj)
    
    Dim nSub As Double
    nSub = 2 ^ (mObj - mBase)
    
    If nSub > 10000 Then
        If MsgBox("Se generaran " & nSub & " subredes. Puede tardar. Continuar?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    
    ws.Range("A4").Value = "ID red": ws.Range("B4").Value = DblToIp(ipD): ws.Range("C4").Value = "Mascara": ws.Range("D4").Value = ConvDec(mBase): ws.Range("E4").Value = "/N": ws.Range("F4").Value = "/" & mBase
    ws.Range("A5").Value = "Bits de red": ws.Range("B5").Value = mBase & " bits": ws.Range("C5").Value = "Bits de host": ws.Range("D5").Value = (32 - mBase) & " bits": ws.Range("E5").Value = "Host utiles": ws.Range("F5").Value = "2^" & (32 - mBase) & "-2=" & (2 ^ (32 - mBase) - 2)
    
    Dim r As Long
    r = 7
    
    Dim hds() As Variant
    hds = Array("Network Address", "Subnet Mask", "/N", "First Usable IP", "Last Usable IP", "Broadcast IP", "Hosts utiles")
    
    Dim i As Integer
    For i = 0 To UBound(hds)
        ws.Cells(r, i + 2).Value = hds(i)
    Next i
    
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Dim c As Double
    For c = 0 To nSub - 1
        r = r + 1
        Dim nD As Double
        nD = ipD + (c * bT)
        Dim bcD As Double
        bcD = nD + bT - 1
        
        ws.Cells(r, 2).Value = DblToIp(nD)
        ws.Cells(r, 3).Value = ConvDec(mObj)
        ws.Cells(r, 4).Value = "/" & mObj
        
        If bT >= 4 Then
            ws.Cells(r, 5).Value = DblToIp(nD + 1)
            ws.Cells(r, 6).Value = DblToIp(bcD - 1)
            ws.Cells(r, 7).Value = DblToIp(bcD)
            ws.Cells(r, 8).Value = bT - 2
        Else
            ws.Cells(r, 5).Value = "-"
            ws.Cells(r, 6).Value = "-"
            ws.Cells(r, 7).Value = DblToIp(bcD)
            ws.Cells(r, 8).Value = 0
        End If
        
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
    Next c
    
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    ws.Range(ws.Cells(7, 2), ws.Cells(r, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
    ws.Range(ws.Cells(7, 2), ws.Cells(r, 2)).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    
    ws.Range("A1:H" & r).ColumnWidth = 15
    ws.Range("A1:H" & r).HorizontalAlignment = xlCenter
    ws.Range("A1:H" & r).VerticalAlignment = xlCenter
    
    On Error Resume Next
    ws.Range("A4:H" & r).Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
    
End Sub

Function ConvDec(p As Integer) As String
    Dim mK As Double: mK = 0
    Dim i As Integer
    For i = 31 To (32 - p) Step -1
        mK = mK + (2 ^ i)
    Next i
    ConvDec = DblToIp(mK)
End Function

Function IpToDbl(ip As String) As Double
    Dim p() As String: p = Split(ip, ".")
    IpToDbl = CDbl(p(0)) * 16777216# + CDbl(p(1)) * 65536# + CDbl(p(2)) * 256# + CDbl(p(3))
End Function

Function DblToIp(ByVal val As Double) As String
    Dim o1 As Long, o2 As Long, o3 As Long, o4 As Long
    o1 = Int(val / 16777216#): val = val - (CDbl(o1) * 16777216#)
    o2 = Int(val / 65536#):    val = val - (CDbl(o2) * 65536#)
    o3 = Int(val / 256#):     o4 = val - (CDbl(o3) * 256#)
    DblToIp = o1 & "." & o2 & "." & o3 & "." & o4
End Function




