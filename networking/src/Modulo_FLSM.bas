Option Explicit

Private Const MAX_FILAS_DETALLE As Long = 128

Sub GenerarFLSM()

    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastR As Long
    lastR = 2
    
    Do While Trim(CStr(ws.Cells(lastR + 1, 1).Value)) <> "" And Trim(CStr(ws.Cells(lastR + 1, 2).Value)) <> ""
        lastR = lastR + 1
    Loop
    
    Dim n As Long
    n = lastR - 2
    
    If n <= 0 Then
        MsgBox "No hay redes para calcular. Ingresa las áreas desde la fila 3.", vbExclamation
        Exit Sub
    End If
    
    Dim ip As String
    Dim ipOriginal As String
    Dim m As Integer
    
    ip = Trim(CStr(ws.Range("A1").Value))
    ipOriginal = ip
    
    If Not EsIPValida(ip) Then
        MsgBox "La IP ingresada en A1 no es válida." & vbCrLf & _
               "Ejemplo correcto: 192.168.1.0", vbCritical
        Exit Sub
    End If
    
    If Not IsNumeric(ws.Range("B1").Value) Then
        MsgBox "La máscara ingresada en B1 no es válida." & vbCrLf & _
               "Ejemplo correcto: 24", vbCritical
        Exit Sub
    End If
    
    If CDbl(ws.Range("B1").Value) <> Fix(CDbl(ws.Range("B1").Value)) Then
        MsgBox "La máscara debe ser un número entero." & vbCrLf & _
               "Ejemplo correcto: 24", vbCritical
        Exit Sub
    End If
    
    m = CInt(ws.Range("B1").Value)
    
    If m < 0 Or m > 30 Then
        MsgBox "La máscara base debe estar entre 0 y 30 para FLSM clásico.", vbCritical
        Exit Sub
    End If
    
    Dim ipDOriginal As Double
    Dim ipDBase As Double
    Dim bcBase As Double
    Dim totalBase As Double
    
    ipDOriginal = IpToDbl(ip)
    ipDBase = NetworkAddressDbl(ipDOriginal, m)
    bcBase = ipDBase + (2 ^ (32 - m)) - 1
    totalBase = 2 ^ (32 - m)
    
    If ipDOriginal <> ipDBase Then
        ip = DblToIp(ipDBase)
    End If
    
    Dim listN() As String
    Dim listH() As Double
    
    ReDim listN(1 To n)
    ReDim listH(1 To n)
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To n
        
        listN(i) = Trim(CStr(ws.Cells(i + 2, 1).Value))
        
        If listN(i) = "" Then
            MsgBox "Falta el nombre del área en la fila " & (i + 2) & ".", vbCritical
            Exit Sub
        End If
        
        If Not IsNumeric(ws.Cells(i + 2, 2).Value) Then
            MsgBox "La cantidad de hosts en la fila " & (i + 2) & " no es válida.", vbCritical
            Exit Sub
        End If
        
        listH(i) = CDbl(ws.Cells(i + 2, 2).Value)
        
        If listH(i) <= 0 Then
            MsgBox "La cantidad de hosts debe ser mayor a cero en la fila " & (i + 2) & ".", vbCritical
            Exit Sub
        End If
        
        If listH(i) <> Fix(listH(i)) Then
            MsgBox "La cantidad de hosts debe ser un número entero en la fila " & (i + 2) & ".", vbCritical
            Exit Sub
        End If
        
        If listH(i) > 4294967294# Then
            MsgBox "La cantidad de hosts en la fila " & (i + 2) & " es demasiado grande para IPv4.", vbCritical
            Exit Sub
        End If
        
    Next i
    
    Dim tH As Double
    Dim tN As String
    
    For i = 1 To n - 1
        For j = i + 1 To n
            If listH(i) < listH(j) Then
                
                tH = listH(i)
                listH(i) = listH(j)
                listH(j) = tH
                
                tN = listN(i)
                listN(i) = listN(j)
                listN(j) = tN
                
            End If
        Next j
    Next i
    
    Dim n1 As Integer
    n1 = BitsParaHosts(listH(1))
    
    Dim m1 As Integer
    m1 = 32 - n1
    
    If m1 < m Then
        MsgBox "FLSM Error: La red con más hosts necesita una máscara /" & m1 & _
               ", pero la red base es /" & m & "." & vbCrLf & _
               "La red base no alcanza.", vbCritical
        Exit Sub
    End If
    
    Dim sb1 As Integer
    sb1 = m1 - m
    
    Dim maxSub As Double
    maxSub = 2 ^ sb1
    
    If maxSub < n Then
        MsgBox "FLSM Error: La máscara /" & m1 & " solo permite " & NumTxt(maxSub) & _
               " subredes, pero solicitaste " & n & ".", vbCritical
        Exit Sub
    End If
    
    Dim blockSize As Double
    blockSize = 2 ^ (32 - m1)
    
    If ipDBase + (CDbl(n) * blockSize) - 1 > bcBase Then
        MsgBox "La red base no alcanza para todas las subredes solicitadas.", vbCritical
        Exit Sub
    End If
    
    Dim oldCalc As XlCalculation
    Dim appTocada As Boolean
    
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    appTocada = True
    
    With ws.Range("A" & (lastR + 1) & ":I" & ws.Rows.Count)
        .UnMerge
        .Clear
    End With
    
    Dim outStart As Long
    outStart = lastR + 3
    
    ws.Range("A" & outStart).Value = "ID red"
    ws.Range("B" & outStart).Value = ip
    ws.Range("C" & outStart).Value = "Mascara"
    ws.Range("D" & outStart).Value = ConvDec(m)
    ws.Range("E" & outStart).Value = "/N"
    ws.Range("F" & outStart).Value = "/" & m
    
    ws.Range("A" & (outStart + 1)).Value = "Bits de red"
    ws.Range("B" & (outStart + 1)).Value = m & " bits"
    ws.Range("C" & (outStart + 1)).Value = "Bits de host"
    ws.Range("D" & (outStart + 1)).Value = (32 - m) & " bits"
    ws.Range("E" & (outStart + 1)).Value = "Host utiles"
    ws.Range("F" & (outStart + 1)).Value = "2^" & (32 - m) & "-2=" & NumTxt((2 ^ (32 - m)) - 2)
    
    If ipOriginal <> ip Then
        ws.Range("A" & (outStart + 2)).Value = "Nota"
        ws.Range("B" & (outStart + 2)).Value = "La IP ingresada era " & ipOriginal & _
                                                " y se ajustó al ID de red correcto: " & ip
        ws.Range("B" & (outStart + 2) & ":F" & (outStart + 2)).Merge
    End If
    
    ws.Range("A" & (outStart + 3)).Value = "Paso 1"
    ws.Range("B" & (outStart + 3)).Value = "Red mas hosts"
    ws.Range("C" & (outStart + 3)).Value = listN(1) & ": " & NumTxt(listH(1)) & " hosts"
    
    ws.Range("A" & (outStart + 4)).Value = "Paso 2"
    ws.Range("B" & (outStart + 4)).Value = "Bits minimos"
    ws.Range("C" & (outStart + 4)).Value = "2^n - 2 >= " & NumTxt(listH(1)) & " hosts"
    ws.Range("D" & (outStart + 4)).Value = "n = " & n1 & " bits host"
    
    ws.Range("A" & (outStart + 5)).Value = "Paso 3"
    ws.Range("B" & (outStart + 5)).Value = "Subnetear red"
    ws.Range("C" & (outStart + 5)).Value = (32 - m) & " bits - " & n1 & " bits = " & sb1 & " bits para subredes"
    ws.Range("C" & (outStart + 5) & ":D" & (outStart + 5)).Merge
    ws.Range("E" & (outStart + 5)).Value = "32 - " & n1 & " = " & m1
    ws.Range("F" & (outStart + 5)).Value = "/" & m1
    
    ws.Range("A" & (outStart + 6)).Value = "Paso 4"
    ws.Range("B" & (outStart + 6)).Value = "Asignar subredes FLSM"
    
    Dim ipD As Double
    ipD = ipDBase
    
    Dim r As Long
    r = outStart + 8
    
    ws.Cells(r, 2).Value = NumTxt(listH(1)) & " hosts (Max)"
    ws.Cells(r, 3).Value = "2^n - 2 >= " & NumTxt(listH(1)) & " hosts"
    ws.Cells(r, 4).Value = "n = " & n1 & " bits"
    ws.Cells(r, 5).Value = sb1 & " bits subredes"
    ws.Cells(r, 6).Value = "/" & m1
    ws.Cells(r, 7).Value = "Total: " & NumTxt(maxSub) & " subredes"
    
    r = r + 1
    
    Dim limitSub As Long
    
    If maxSub > MAX_FILAS_DETALLE Then
        limitSub = MAX_FILAS_DETALLE
    Else
        limitSub = CLng(maxSub)
    End If
    
    Dim c As Long
    
    For c = 0 To limitSub - 1
        
        Dim rowD As Double
        rowD = ipD + CDbl(c) * blockSize
        
        Dim o1 As Long
        Dim o2 As Long
        Dim o3 As Long
        Dim o4 As Long
        Dim tmp As Double
        
        tmp = rowD
        
        o1 = Fix(tmp / 16777216#)
        tmp = tmp - (CDbl(o1) * 16777216#)
        
        o2 = Fix(tmp / 65536#)
        tmp = tmp - (CDbl(o2) * 65536#)
        
        o3 = Fix(tmp / 256#)
        o4 = CLng(tmp - (CDbl(o3) * 256#))
        
        Call Pinta(ws.Cells(r, 2), o1, 1, m, m1)
        Call Pinta(ws.Cells(r, 3), o2, 2, m, m1)
        Call Pinta(ws.Cells(r, 4), o3, 3, m, m1)
        Call Pinta(ws.Cells(r, 5), o4, 4, m, m1)
        
        ws.Cells(r, 6).Value = DblToIp(rowD)
        ws.Cells(r, 7).Value = "/" & m1
        
        If c < n Then
            
            ws.Cells(r, 8).Value = listN(c + 1)
            ws.Cells(r, 8).Font.Color = RGB(0, 0, 0)
            ws.Cells(r, 8).Font.Bold = False
            
            With ws.Range(ws.Cells(r, 2), ws.Cells(r, 8))
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                
                .Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
                .Borders(xlEdgeTop).Color = RGB(0, 0, 0)
                .Borders(xlEdgeRight).Color = RGB(0, 0, 0)
                .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
            End With
            
        Else
            
            ws.Cells(r, 8).Value = "Disponible"
            ws.Cells(r, 8).Font.Color = RGB(128, 128, 128)
            ws.Cells(r, 8).Font.Bold = False
            
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.LineStyle = xlContinuous
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.Color = RGB(220, 220, 220)
            
        End If
        
        r = r + 1
        
    Next c
    
    If maxSub > MAX_FILAS_DETALLE Then
        ws.Cells(r, 2).Value = "..."
        ws.Cells(r, 6).Value = "Se omitieron " & NumTxt(maxSub - MAX_FILAS_DETALLE) & _
                               " subredes visuales para evitar que Excel se congele"
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.Color = RGB(220, 220, 220)
        r = r + 1
    End If
    
    r = r + 1
    
    Dim sh As Long
    sh = r
    
    Dim hds As Variant
    hds = Array("Area", "ID de red", "Mascara /N", "Mascara decimal", "Primera IP", "Ultima IP", "IP Broadcast")
    
    For i = 0 To UBound(hds)
        ws.Cells(r, i + 2).Value = hds(i)
    Next i
    
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
    Dim k As Long
    
    For k = 1 To n
        
        r = r + 1
        
        Dim idD As Double
        Dim bcD As Double
        
        idD = ipD + CDbl(k - 1) * blockSize
        bcD = idD + blockSize - 1
        
        ws.Cells(r, 2).Value = listN(k)
        ws.Cells(r, 3).Value = DblToIp(idD)
        ws.Cells(r, 4).Value = "/" & m1
        ws.Cells(r, 5).Value = ConvDec(m1)
        ws.Cells(r, 6).Value = DblToIp(idD + 1)
        ws.Cells(r, 7).Value = DblToIp(bcD - 1)
        ws.Cells(r, 8).Value = DblToIp(bcD)
        
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders(xlEdgeBottom).Color = RGB(220, 220, 220)
        
    Next k
    
    ws.Range(ws.Cells(sh, 2), ws.Cells(r, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
    ws.Range(ws.Cells(sh, 2), ws.Cells(r, 2)).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    
    ws.Range("A1:H" & r).ColumnWidth = 15
    ws.Range("A1:H" & r).HorizontalAlignment = xlCenter
    ws.Range("A1:H" & r).VerticalAlignment = xlCenter
    
    ws.Range("B" & (outStart + 3) & ":C" & (outStart + 6)).HorizontalAlignment = xlLeft
    
    On Error Resume Next
    ws.Range("A" & outStart & ":I" & r).Errors(xlNumberAsText).Ignore = True
    On Error GoTo ErrHandler
    
    MsgBox "Cálculo FLSM generado correctamente.", vbInformation

Finalizar:
    
    If appTocada Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = oldCalc
    End If
    
    Exit Sub

ErrHandler:
    
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
    Resume Finalizar

End Sub

Function ConvDec(ByVal p As Integer) As String
    
    If p < 0 Or p > 32 Then
        ConvDec = ""
        Exit Function
    End If
    
    Dim mK As Double
    mK = 0
    
    Dim i As Integer
    
    For i = 31 To (32 - p) Step -1
        mK = mK + (2 ^ i)
    Next i
    
    ConvDec = DblToIp(mK)
    
End Function

Function IpToDbl(ByVal ip As String) As Double
    
    Dim p() As String
    p = Split(Trim(CStr(ip)), ".")
    
    IpToDbl = CDbl(p(0)) * 16777216# + _
              CDbl(p(1)) * 65536# + _
              CDbl(p(2)) * 256# + _
              CDbl(p(3))
              
End Function

Function DblToIp(ByVal val As Double) As String
    
    Dim o1 As Long
    Dim o2 As Long
    Dim o3 As Long
    Dim o4 As Long
    
    If val < 0 Or val > 4294967295# Then
        DblToIp = "FUERA_RANGO"
        Exit Function
    End If
    
    o1 = Fix(val / 16777216#)
    val = val - (CDbl(o1) * 16777216#)
    
    o2 = Fix(val / 65536#)
    val = val - (CDbl(o2) * 65536#)
    
    o3 = Fix(val / 256#)
    o4 = CLng(val - (CDbl(o3) * 256#))
    
    DblToIp = o1 & "." & o2 & "." & o3 & "." & o4
    
End Function

Function EsIPValida(ByVal ip As String) As Boolean
    
    Dim p() As String
    Dim i As Long
    Dim oct As Long
    
    ip = Trim(CStr(ip))
    p = Split(ip, ".")
    
    If UBound(p) <> 3 Then
        EsIPValida = False
        Exit Function
    End If
    
    For i = 0 To 3
        
        If Not SoloDigitos(p(i)) Then
            EsIPValida = False
            Exit Function
        End If
        
        oct = CLng(p(i))
        
        If oct < 0 Or oct > 255 Then
            EsIPValida = False
            Exit Function
        End If
        
    Next i
    
    EsIPValida = True
    
End Function

Function SoloDigitos(ByVal txt As String) As Boolean
    
    Dim i As Long
    Dim ch As String
    
    If Len(txt) = 0 Then
        SoloDigitos = False
        Exit Function
    End If
    
    For i = 1 To Len(txt)
        
        ch = Mid$(txt, i, 1)
        
        If ch < "0" Or ch > "9" Then
            SoloDigitos = False
            Exit Function
        End If
        
    Next i
    
    SoloDigitos = True
    
End Function

Function NetworkAddressDbl(ByVal ipD As Double, ByVal prefijo As Integer) As Double
    
    Dim bloque As Double
    bloque = 2 ^ (32 - prefijo)
    
    NetworkAddressDbl = Fix(ipD / bloque) * bloque
    
End Function

Function BitsParaHosts(ByVal hosts As Double) As Integer
    
    Dim n As Integer
    n = 0
    
    Do While (2 ^ n) - 2 < hosts
        
        n = n + 1
        
        If n > 32 Then
            Err.Raise vbObjectError + 1000, , "La cantidad de hosts excede la capacidad de IPv4."
        End If
        
    Loop
    
    BitsParaHosts = n
    
End Function

Function NumTxt(ByVal v As Double) As String
    
    NumTxt = Format$(v, "0")
    
End Function

Sub Pinta(ByVal cell As Range, ByVal v As Long, ByVal num As Integer, ByVal p1 As Integer, ByVal p2 As Integer)
    
    Dim s As String
    s = ""
    
    Dim i As Integer
    Dim g As Integer
    Dim idx As Integer
    
    For i = 7 To 0 Step -1
        
        If (v And CLng(2 ^ i)) <> 0 Then
            s = s & "1"
        Else
            s = s & "0"
        End If
        
        If i = 4 Then s = s & " "
        
    Next i
    
    If num < 4 Then s = s & "."
    
    cell.NumberFormat = "@"
    cell.Value = s
    
    On Error Resume Next
    cell.Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
    
    cell.Font.Color = RGB(68, 114, 196)
    
    For i = 7 To 0 Step -1
        
        g = (num - 1) * 8 + (8 - i)
        
        If i >= 4 Then
            idx = 8 - i
        Else
            idx = 9 - i
        End If
        
        If g > p1 And g <= p2 Then
            cell.Characters(Start:=idx, Length:=1).Font.Color = RGB(255, 0, 0)
        ElseIf g > p2 Then
            cell.Characters(Start:=idx, Length:=1).Font.Color = RGB(0, 0, 0)
        End If
        
    Next i
    
End Sub



