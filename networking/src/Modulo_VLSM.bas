Option Explicit

Private Const MAX_FILAS_DETALLE As Long = 512
Private Const MAX_BLOQUES_LIBRES As Long = 20000

Sub GenerarVLSM()

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
        MsgBox "La máscara base debe estar entre 0 y 30 para VLSM clásico.", vbCritical
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
    
    Dim sumaBloques As Double
    Dim bitTemp As Integer
    Dim prefTemp As Integer
    
    sumaBloques = 0
    
    For i = 1 To n
        
        bitTemp = BitsParaHosts(listH(i))
        prefTemp = 32 - bitTemp
        
        If prefTemp < m Then
            MsgBox "La red '" & listN(i) & "' necesita una subred /" & prefTemp & _
                   ", pero la red base es /" & m & "." & vbCrLf & _
                   "La red base no alcanza para esa cantidad de hosts.", vbCritical
            Exit Sub
        End If
        
        sumaBloques = sumaBloques + (2 ^ bitTemp)
        
    Next i
    
    If sumaBloques > totalBase Then
        MsgBox "La red base no alcanza para todas las subredes solicitadas." & vbCrLf & vbCrLf & _
               "Direcciones disponibles en la red base: " & NumTxt(totalBase) & vbCrLf & _
               "Direcciones requeridas por VLSM: " & NumTxt(sumaBloques), vbCritical
        Exit Sub
    End If
    
    Dim resN() As String
    Dim resID() As Double
    Dim resP() As Integer
    Dim parentID() As Double
    Dim parentP() As Integer
    
    ReDim resN(1 To n)
    ReDim resID(1 To n)
    ReDim resP(1 To n)
    ReDim parentID(1 To n)
    ReDim parentP(1 To n)
    
    Dim freeStart() As Double
    Dim freeP() As Integer
    Dim freeCount As Long
    
    ReDim freeStart(1 To 1)
    ReDim freeP(1 To 1)
    
    freeCount = 1
    freeStart(1) = ipDBase
    freeP(1) = m
    
    Dim k As Long
    
    For k = 1 To n
        
        Dim h As Double
        h = listH(k)
        
        Dim bitN As Integer
        bitN = BitsParaHosts(h)
        
        Dim nP As Integer
        nP = 32 - bitN
        
        Dim idxFree As Long
        idxFree = BuscarBloqueLibre(freeStart, freeP, freeCount, nP)
        
        If idxFree = 0 Then
            MsgBox "No se encontró bloque disponible para: " & listN(k), vbCritical
            Exit Sub
        End If
        
        Dim pPadre As Integer
        Dim idPadre As Double
        
        idPadre = freeStart(idxFree)
        pPadre = freeP(idxFree)
        
        parentID(k) = idPadre
        parentP(k) = pPadre
        
        resN(k) = listN(k)
        resID(k) = idPadre
        resP(k) = nP
        
        Call QuitarBloqueLibre(freeStart, freeP, freeCount, idxFree)
        
        Dim blockSize As Double
        blockSize = 2 ^ (32 - nP)
        
        Dim cantidadSubBloques As Double
        cantidadSubBloques = 2 ^ (nP - pPadre)
        
        If cantidadSubBloques - 1 + freeCount > MAX_BLOQUES_LIBRES Then
            MsgBox "La explicación visual generaría demasiados bloques libres." & vbCrLf & _
                   "Usa una red base más pequeña o menos diferencia entre máscaras.", vbCritical
            Exit Sub
        End If
        
        Dim c As Long
        
        For c = 1 To CLng(cantidadSubBloques - 1)
            Call AgregarBloqueLibre(freeStart, freeP, freeCount, idPadre + CDbl(c) * blockSize, nP)
        Next c
        
        Call OrdenarBloquesLibres(freeStart, freeP, freeCount)
        
    Next k
    
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
    
    Dim n1 As Integer
    n1 = BitsParaHosts(listH(1))
    
    Dim m1 As Integer
    m1 = 32 - n1
    
    Dim sb1 As Integer
    sb1 = m1 - m
    
    Dim outStart As Long
    outStart = lastR + 3
    
    ws.Range("A" & outStart).Value = "ID red"
    Call PonerTextoSinError(ws.Range("B" & outStart), ip)
    ws.Range("C" & outStart).Value = "Mascara"
    Call PonerTextoSinError(ws.Range("D" & outStart), ConvDec(m))
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
    ws.Range("B" & (outStart + 6)).Value = "Asignar subredes y repetir"
    
    Dim r As Long
    r = outStart + 8
    
    For k = 1 To n
        
        h = listH(k)
        bitN = BitsParaHosts(h)
        nP = resP(k)
        
        Dim pActualPadre As Integer
        Dim idActualPadre As Double
        
        pActualPadre = parentP(k)
        idActualPadre = parentID(k)
        
        Dim bBits As Integer
        bBits = nP - pActualPadre
        
        Dim rowsTotal As Double
        rowsTotal = 2 ^ bBits
        
        Dim rowsMostrar As Long
        
        If rowsTotal > MAX_FILAS_DETALLE Then
            rowsMostrar = MAX_FILAS_DETALLE
        Else
            rowsMostrar = CLng(rowsTotal)
        End If
        
        ws.Cells(r, 2).Value = NumTxt(h) & " hosts"
        ws.Cells(r, 3).Value = "2^n - 2 >= " & NumTxt(h) & " hosts"
        ws.Cells(r, 4).Value = "n = " & bitN & " bits"
        ws.Cells(r, 5).Value = bBits & " bits subredes"
        ws.Cells(r, 6).Value = "/" & nP
        
        r = r + 1
        
        blockSize = 2 ^ (32 - nP)
        
        For c = 0 To rowsMostrar - 1
            
            Dim rowD As Double
            rowD = idActualPadre + CDbl(c) * blockSize
            
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
            
            Call Pinta(ws.Cells(r, 2), o1, 1, pActualPadre, nP)
            Call Pinta(ws.Cells(r, 3), o2, 2, pActualPadre, nP)
            Call Pinta(ws.Cells(r, 4), o3, 3, pActualPadre, nP)
            Call Pinta(ws.Cells(r, 5), o4, 4, pActualPadre, nP)
            
            Call PonerTextoSinError(ws.Cells(r, 6), DblToIp(rowD))
            ws.Cells(r, 7).Value = "/" & nP
            
            If rowD = resID(k) Then
                
                ws.Cells(r, 8).Value = resN(k)
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
                
                Dim futuro As String
                futuro = BuscarAreaFutura(rowD, blockSize, k, resID, resP, resN, n)
                
                If futuro <> "" Then
                    ws.Cells(r, 8).Value = "Subnet " & futuro
                Else
                    ws.Cells(r, 8).Value = "Disponible"
                End If
                
                ws.Cells(r, 8).Font.Color = RGB(128, 128, 128)
                ws.Cells(r, 8).Font.Bold = False
                
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.LineStyle = xlContinuous
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.Color = RGB(220, 220, 220)
                
            End If
            
            r = r + 1
            
        Next c
        
        If rowsTotal > rowsMostrar Then
            ws.Cells(r, 2).Value = "..."
            ws.Cells(r, 6).Value = "Se omitieron " & NumTxt(rowsTotal - rowsMostrar) & _
                                   " filas para evitar que Excel se congele"
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.LineStyle = xlContinuous
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 8)).Borders.Color = RGB(220, 220, 220)
            r = r + 1
        End If
        
        r = r + 1
        
    Next k
    
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
    
    For k = 1 To n
        
        r = r + 1
        
        Dim idD As Double
        Dim pfx As Integer
        Dim bcD As Double
        
        idD = resID(k)
        pfx = resP(k)
        bcD = idD + (2 ^ (32 - pfx)) - 1
        
        ws.Cells(r, 2).Value = resN(k)
        Call PonerTextoSinError(ws.Cells(r, 3), DblToIp(idD))
        ws.Cells(r, 4).Value = "/" & pfx
        Call PonerTextoSinError(ws.Cells(r, 5), ConvDec(pfx))
        Call PonerTextoSinError(ws.Cells(r, 6), DblToIp(idD + 1))
        Call PonerTextoSinError(ws.Cells(r, 7), DblToIp(bcD - 1))
        Call PonerTextoSinError(ws.Cells(r, 8), DblToIp(bcD))
        
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
    
    MsgBox "Cálculo VLSM generado correctamente.", vbInformation

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

Function BuscarBloqueLibre(ByRef freeStart() As Double, ByRef freeP() As Integer, ByVal freeCount As Long, ByVal prefijoNecesario As Integer) As Long
    
    Dim i As Long
    
    For i = 1 To freeCount
        If freeP(i) <= prefijoNecesario Then
            BuscarBloqueLibre = i
            Exit Function
        End If
    Next i
    
    BuscarBloqueLibre = 0
    
End Function

Sub AgregarBloqueLibre(ByRef freeStart() As Double, ByRef freeP() As Integer, ByRef freeCount As Long, ByVal inicio As Double, ByVal prefijo As Integer)
    
    freeCount = freeCount + 1
    
    ReDim Preserve freeStart(1 To freeCount)
    ReDim Preserve freeP(1 To freeCount)
    
    freeStart(freeCount) = inicio
    freeP(freeCount) = prefijo
    
End Sub

Sub QuitarBloqueLibre(ByRef freeStart() As Double, ByRef freeP() As Integer, ByRef freeCount As Long, ByVal idx As Long)
    
    Dim i As Long
    
    If freeCount <= 1 Then
        freeCount = 0
        ReDim freeStart(1 To 1)
        ReDim freeP(1 To 1)
        Exit Sub
    End If
    
    For i = idx To freeCount - 1
        freeStart(i) = freeStart(i + 1)
        freeP(i) = freeP(i + 1)
    Next i
    
    freeCount = freeCount - 1
    
    ReDim Preserve freeStart(1 To freeCount)
    ReDim Preserve freeP(1 To freeCount)
    
End Sub

Sub OrdenarBloquesLibres(ByRef freeStart() As Double, ByRef freeP() As Integer, ByVal freeCount As Long)
    
    Dim i As Long
    Dim j As Long
    Dim tD As Double
    Dim tP As Integer
    
    If freeCount <= 1 Then Exit Sub
    
    For i = 1 To freeCount - 1
        For j = i + 1 To freeCount
            
            If freeStart(i) > freeStart(j) Then
                
                tD = freeStart(i)
                freeStart(i) = freeStart(j)
                freeStart(j) = tD
                
                tP = freeP(i)
                freeP(i) = freeP(j)
                freeP(j) = tP
                
            End If
            
        Next j
    Next i
    
End Sub

Function BuscarAreaFutura(ByVal inicioBloque As Double, ByVal tamBloque As Double, ByVal kActual As Long, _
                          ByRef resID() As Double, ByRef resP() As Integer, ByRef resN() As String, ByVal total As Long) As String
    
    Dim q As Long
    Dim finBloque As Double
    
    finBloque = inicioBloque + tamBloque - 1
    
    For q = kActual + 1 To total
        
        If resID(q) >= inicioBloque And resID(q) <= finBloque Then
            BuscarAreaFutura = resN(q)
            Exit Function
        End If
        
    Next q
    
    BuscarAreaFutura = ""
    
End Function

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

Sub PonerTextoSinError(ByVal celda As Range, ByVal valor As String)
    
    celda.NumberFormat = "@"
    celda.Value = CStr(valor)
    
    On Error Resume Next
    celda.Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
    
End Sub

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


