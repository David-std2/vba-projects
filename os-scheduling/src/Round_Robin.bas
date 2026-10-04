Option Explicit

' helpers
Private Function CollectionContains(col As Collection, val As Variant) As Boolean
    Dim i As Long
    On Error GoTo ErrHandler
    For i = 1 To col.Count
        If CStr(col(i)) = CStr(val) Then
            CollectionContains = True
            Exit Function
        End If
    Next i
ErrHandler:
    CollectionContains = False
End Function

Private Function IndexOfName(namesArr As Variant, key As String) As Long
    Dim i As Long
    For i = LBound(namesArr) To UBound(namesArr)
        If CStr(namesArr(i)) = CStr(key) Then
            IndexOfName = i - LBound(namesArr) + 1
            Exit Function
        End If
    Next i
    IndexOfName = 0
End Function

Private Function PickColor(idx As Long) As Long
    Dim palette(0 To 7) As Long
    palette(0) = RGB(255, 250, 205)
    palette(1) = RGB(221, 235, 247)
    palette(2) = RGB(255, 230, 204)
    palette(3) = RGB(226, 239, 218)
    palette(4) = RGB(244, 204, 204)
    palette(5) = RGB(215, 228, 188)
    palette(6) = RGB(226, 239, 249)
    palette(7) = RGB(255, 229, 153)
    PickColor = palette((idx - 1) Mod (UBound(palette) + 1))
End Function

' MAIN
Sub Round_Robin_Alg()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' leer tabla A1:C?
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No hay datos en A1:C?. Coloca la tabla (headers en fila 1).", vbExclamation
        Exit Sub
    End If
    
    Dim n As Long
    n = lastRow - 1
    
    Dim namesArr() As Variant
    Dim arrivalsArr() As Long
    Dim burstsArr() As Long
    ReDim namesArr(1 To n)
    ReDim arrivalsArr(1 To n)
    ReDim burstsArr(1 To n)
    
    Dim i As Long
    For i = 1 To n
        namesArr(i) = CStr(ws.Cells(i + 1, 1).Value)
        arrivalsArr(i) = CLng(ws.Cells(i + 1, 2).Value)
        burstsArr(i) = CLng(ws.Cells(i + 1, 3).Value)
    Next i
    
    ' pedir quantum
    Dim qIn As Variant, Q As Long
    qIn = InputBox("Quantum (Q). Enter para usar 4:", "Quantum", 4)
    If Trim(CStr(qIn)) = "" Then Q = 4 Else Q = CLng(qIn)
    If Q <= 0 Then Q = 4
    
    Dim remaining() As Long
    Dim isFinished() As Boolean
    Dim finish_time() As Long
    ReDim remaining(1 To n)
    ReDim isFinished(1 To n)
    ReDim finish_time(1 To n)
    
    For i = 1 To n
        remaining(i) = burstsArr(i)
        isFinished(i) = False
        finish_time(i) = -1
    Next i
    
    Dim queue As Collection: Set queue = New Collection
    Dim toRequeue As Collection: Set toRequeue = New Collection
    Dim cpu As String: cpu = ""
    Dim cpu_qleft As Long: cpu_qleft = 0
    Dim t As Long: t = 0
    Dim maxCols As Long: maxCols = 4000
    Dim tableState() As Variant
    ReDim tableState(1 To n, 1 To maxCols)
    
    Dim allDone As Boolean, idxProc As Long, j As Long, k As Long
    
    ' Round-Robin
    Do
        allDone = True
        For i = 1 To n
            If Not isFinished(i) Then
                allDone = False
                Exit For
            End If
        Next i
        If allDone Then Exit Do
        
        If toRequeue.Count > 0 Then
            For k = 1 To toRequeue.Count
                If Not CollectionContains(queue, toRequeue(k)) Then queue.Add toRequeue(k)
            Next k
            Set toRequeue = New Collection
        End If
        
        For i = 1 To n
            If arrivalsArr(i) = t Then
                If Not CollectionContains(queue, namesArr(i)) And Not isFinished(i) Then queue.Add namesArr(i)
            End If
        Next i
        
        If cpu = "" Then
            If queue.Count > 0 Then
                cpu = queue(1)
                queue.Remove 1
                idxProc = IndexOfName(namesArr, cpu)
                If idxProc > 0 Then
                    If remaining(idxProc) < Q Then
                        cpu_qleft = remaining(idxProc)
                    Else
                        cpu_qleft = Q
                    End If
                End If
            End If
        End If
        
        For i = 1 To n
            tableState(i, t + 1) = ""
        Next i
        
        If cpu <> "" Then
            idxProc = IndexOfName(namesArr, cpu)
            If idxProc > 0 Then tableState(idxProc, t + 1) = "0"
        End If
        
        If queue.Count > 0 Then
            For j = 1 To queue.Count
                Dim qname As String: qname = queue(j)
                idxProc = IndexOfName(namesArr, qname)
                If idxProc > 0 Then tableState(idxProc, t + 1) = CStr(j)
            Next j
        End If
        
        If cpu <> "" Then
            idxProc = IndexOfName(namesArr, cpu)
            If idxProc > 0 Then
                remaining(idxProc) = remaining(idxProc) - 1
                cpu_qleft = cpu_qleft - 1
                If remaining(idxProc) = 0 Then
                    isFinished(idxProc) = True
                    finish_time(idxProc) = t + 1  ' Fi en t+1
                    cpu = ""
                    cpu_qleft = 0
                ElseIf cpu_qleft = 0 Then
                    toRequeue.Add namesArr(idxProc)
                    cpu = ""
                End If
            End If
        End If
        
        t = t + 1
        If t >= maxCols - 2 Then Exit Do
    Loop
    
    Dim last As Long: last = 0
    For i = 1 To n
        If finish_time(i) > last Then last = finish_time(i)
    Next i
    If last = 0 Then last = t
    
    Call DibujarDiagramas(ws, namesArr, arrivalsArr, burstsArr, tableState, finish_time, last)
    Call DibujarMetricas(ws, namesArr, arrivalsArr, burstsArr, finish_time, n, last)
    
    MsgBox "Round-Robin generado (orden de reencolado corregido).", vbInformation
End Sub

' dibujar diagramas
Private Sub DibujarDiagramas(ws As Worksheet, namesArr() As Variant, arrivalsArr() As Long, burstsArr() As Long, tableState() As Variant, finish_time() As Long, last As Long)
    Dim n As Long
    n = UBound(namesArr)
    
    Dim baseRow As Long: baseRow = 10
    Dim baseCol As Long: baseCol = 5
    
    Dim clearHeight As Long: clearHeight = 4 * (n + 8) + 60
    ws.Range(ws.Cells(baseRow - 3, baseCol - 1), ws.Cells(baseRow - 3 + clearHeight, baseCol + last + 6)).Clear
    
    Dim headerGrey As Long: headerGrey = RGB(230, 230, 230)
    Dim colGrey1 As Long: colGrey1 = RGB(245, 245, 245)
    Dim colGrey2 As Long: colGrey2 = RGB(235, 235, 235)
    Dim pastel0 As Long: pastel0 = RGB(255, 255, 153)
    
    Dim i As Long, j As Long
    Dim areaArr As Range, areaProc As Range, areaG As Range, areaCompact As Range
    
    ' Diagrama 1: Llegadas
    Dim r0 As Long: r0 = baseRow
    ws.Cells(r0 - 1, baseCol).Value = "Llegadas"
    ws.Cells(r0, baseCol).Value = "Process"
    ws.Cells(r0, baseCol).Font.Bold = True
    
    For j = 0 To last
        ws.Cells(r0, baseCol + j + 1).Value = j
        ws.Cells(r0, baseCol + j + 1).Interior.Color = headerGrey
        ws.Cells(r0, baseCol + j + 1).HorizontalAlignment = xlCenter
    Next j
    
    For i = 1 To n
        ws.Cells(r0 + i, baseCol).Value = namesArr(i)
        ws.Cells(r0 + i, baseCol).Interior.Color = colGrey1
        For j = 0 To last
            If arrivalsArr(i) = j Then
                ws.Cells(r0 + i, baseCol + j + 1).Value = "L"
                ws.Cells(r0 + i, baseCol + j + 1).HorizontalAlignment = xlCenter
            Else
                ws.Cells(r0 + i, baseCol + j + 1).Value = ""
            End If
            If tableState(i, j + 1) = "0" Then
                ws.Cells(r0 + i, baseCol + j + 1).Interior.Color = pastel0
            End If
        Next j
    Next i
    
    Set areaArr = ws.Range(ws.Cells(r0, baseCol), ws.Cells(r0 + n, baseCol + last + 1))
    With areaArr.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Diagrama 2: Procesos
    Dim r1 As Long: r1 = r0 + n + 4
    ws.Cells(r1 - 1, baseCol).Value = "Procesos en Cola"
    ws.Cells(r1, baseCol).Value = "Process": ws.Cells(r1, baseCol).Font.Bold = True
    
    For j = 0 To last
        ws.Cells(r1, baseCol + j + 1).Value = j
        ws.Cells(r1, baseCol + j + 1).Interior.Color = headerGrey
        ws.Cells(r1, baseCol + j + 1).HorizontalAlignment = xlCenter
    Next j
    
    For i = 1 To n
        ws.Cells(r1 + i, baseCol).Value = namesArr(i)
        ws.Cells(r1 + i, baseCol).Interior.Color = colGrey1
        For j = 0 To last
            Dim vv As Variant
            vv = tableState(i, j + 1)
            If finish_time(i) = j Then
                ws.Cells(r1 + i, baseCol + j + 1).Value = "Fi"
                ws.Cells(r1 + i, baseCol + j + 1).Font.Color = RGB(200, 0, 0)
                ws.Cells(r1 + i, baseCol + j + 1).HorizontalAlignment = xlCenter
            ElseIf vv <> "" Then
                ws.Cells(r1 + i, baseCol + j + 1).Value = vv
                If vv = "0" Then
                    ws.Cells(r1 + i, baseCol + j + 1).Interior.Color = pastel0
                End If
                ws.Cells(r1 + i, baseCol + j + 1).HorizontalAlignment = xlCenter
            Else
                ws.Cells(r1 + i, baseCol + j + 1).Value = ""
            End If
        Next j
    Next i
    
    Set areaProc = ws.Range(ws.Cells(r1, baseCol), ws.Cells(r1 + n, baseCol + last + 1))
    With areaProc.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Diagrama 3: Gantt
    Dim r2 As Long: r2 = r1 + n + 4
    ws.Cells(r2 - 1, baseCol).Value = "Diagrama de Gantt"
    ws.Cells(r2, baseCol).Value = "Process": ws.Cells(r2, baseCol).Font.Bold = True
    
    For j = 0 To last
        ws.Cells(r2, baseCol + j + 1).Value = j
        ws.Cells(r2, baseCol + j + 1).Interior.Color = headerGrey
        ws.Cells(r2, baseCol + j + 1).HorizontalAlignment = xlCenter
    Next j
    For i = 1 To n
        ws.Cells(r2 + i, baseCol).Value = namesArr(i)
        ws.Cells(r2 + i, baseCol).Interior.Color = colGrey1
    Next i
    
    ' cpu_at_time
    Dim cpu_at_time() As String
    ReDim cpu_at_time(0 To last)
    For j = 0 To last
        cpu_at_time(j) = ""
        For i = 1 To n
            If tableState(i, j + 1) = "0" Then
                cpu_at_time(j) = namesArr(i)
                Exit For
            End If
        Next i
    Next j
    
    For j = 0 To last
        Dim altColor As Long
        If j Mod 2 = 0 Then altColor = colGrey2 Else altColor = colGrey1
        ws.Range(ws.Cells(r2 + 1, baseCol + j + 1), ws.Cells(r2 + n, baseCol + j + 1)).Interior.Color = altColor
    Next j
    
    Dim segStart As Long, segName As String
    For i = 1 To n
        segStart = -1
        segName = ""
        For j = 0 To last
            If cpu_at_time(j) = namesArr(i) Then
                If segStart = -1 Then segStart = j
                segName = namesArr(i)
            Else
                If segStart <> -1 Then
                    ws.Range(ws.Cells(r2 + i, baseCol + segStart + 1), ws.Cells(r2 + i, baseCol + j)).Merge
                    ws.Cells(r2 + i, baseCol + segStart + 1).Value = segName
                    ws.Cells(r2 + i, baseCol + segStart + 1).HorizontalAlignment = xlCenter
                    ws.Cells(r2 + i, baseCol + segStart + 1).Interior.Color = PickColor(i)
                    segStart = -1
                End If
            End If
        Next j
        If segStart <> -1 Then
            ws.Range(ws.Cells(r2 + i, baseCol + segStart + 1), ws.Cells(r2 + i, baseCol + last + 1)).Merge
            ws.Cells(r2 + i, baseCol + segStart + 1).Value = segName
            ws.Cells(r2 + i, baseCol + segStart + 1).HorizontalAlignment = xlCenter
            ws.Cells(r2 + i, baseCol + segStart + 1).Interior.Color = PickColor(i)
        End If
    Next i
    
    Set areaG = ws.Range(ws.Cells(r2, baseCol), ws.Cells(r2 + n, baseCol + last + 1))
    With areaG.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Diagrama 4: Gantt compacto
    Dim r3 As Long: r3 = r2 + n + 4
    ws.Cells(r3 - 1, baseCol).Value = "Gantt (secuencial)"
    ws.Cells(r3, baseCol).Value = "Process": ws.Cells(r3, baseCol).Font.Bold = True
    
    Dim runStart As Long: runStart = -1
    Dim runName As String: runName = ""
    For j = 0 To last
        If cpu_at_time(j) <> runName Then
            If runName <> "" Then
                ws.Range(ws.Cells(r3, baseCol + runStart + 1), ws.Cells(r3, baseCol + j)).Merge
                ws.Cells(r3, baseCol + runStart + 1).Value = runName
                ws.Cells(r3, baseCol + runStart + 1).HorizontalAlignment = xlCenter
                ws.Cells(r3, baseCol + runStart + 1).Interior.Color = PickColor(IndexOfName(namesArr, runName))
            End If
            runName = cpu_at_time(j)
            runStart = j
        End If
    Next j
    If runName <> "" Then
        ws.Range(ws.Cells(r3, baseCol + runStart + 1), ws.Cells(r3, baseCol + last + 1)).Merge
        ws.Cells(r3, baseCol + runStart + 1).Value = runName
        ws.Cells(r3, baseCol + runStart + 1).HorizontalAlignment = xlCenter
        ws.Cells(r3, baseCol + runStart + 1).Interior.Color = PickColor(IndexOfName(namesArr, runName))
    End If
    
    For j = 0 To last
        ws.Cells(r3 + 1, baseCol + j + 1).Value = j
    Next j

    With ws.Range(ws.Cells(r3 + 1, baseCol + 1), ws.Cells(r3 + 1, baseCol + last + 1))
        .HorizontalAlignment = xlCenter
    End With
    
    Set areaCompact = ws.Range(ws.Cells(r3, baseCol), ws.Cells(r3 + 1, baseCol + last + 1))
    With areaCompact.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Dim colIndex As Long
    For colIndex = baseCol + 1 To baseCol + last + 1
        ws.Columns(colIndex).ColumnWidth = 3
    Next colIndex
End Sub

' dibujar metricas WT y CT
Private Sub DibujarMetricas(ws As Worksheet, namesArr() As Variant, arrivalsArr() As Long, burstsArr() As Long, finish_time() As Long, n As Long, last As Long)
    Dim baseCol As Long: baseCol = 5
    Dim diagramsRightMost As Long
    diagramsRightMost = baseCol + last + 1
    Dim gap As Long: gap = 1
    Dim startCol As Long
    startCol = diagramsRightMost + gap + 1
    
    Dim startRow As Long: startRow = 10
    
    ws.Cells(startRow - 1, startCol).Value = "WT y CT"
    ws.Cells(startRow - 1, startCol).HorizontalAlignment = xlLeft

    ws.Cells(startRow, startCol).Value = "Process"
    ws.Cells(startRow, startCol + 1).Value = "(WT) Waiting Time"
    ws.Cells(startRow, startCol + 2).Value = "(CT) Complete Time"
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 2)).Font.Bold = True
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 2)).Interior.Color = RGB(230, 230, 230)
    
    Dim i As Long
    Dim WT As Double, CT As Double
    Dim sumWT As Double: sumWT = 0
    Dim sumCT As Double: sumCT = 0
    
    For i = 1 To n
        CT = finish_time(i) - arrivalsArr(i)
        WT = CT - burstsArr(i)
        ws.Cells(startRow + i, startCol).Value = namesArr(i)
        ws.Cells(startRow + i, startCol + 1).Value = WT
        ws.Cells(startRow + i, startCol + 2).Value = CT
        
        sumWT = sumWT + WT
        sumCT = sumCT + CT
    Next i
    
    Dim areaMetrics As Range
    Set areaMetrics = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + n, startCol + 2))
    With areaMetrics.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    areaMetrics.HorizontalAlignment = xlCenter
    areaMetrics.VerticalAlignment = xlCenter
    
    Dim sepRow As Long
    sepRow = startRow + n + 3
    
    ws.Cells(sepRow, startCol).Value = "Promedios"
    ws.Cells(sepRow, startCol).Font.Bold = False
    ws.Range(ws.Cells(sepRow, startCol), ws.Cells(sepRow, startCol + 2)).Interior.Color = xlNone
    
    Dim headerGrey As Long: headerGrey = RGB(230, 230, 230)
    ws.Cells(sepRow + 1, startCol + 1).Value = "AWT"
    ws.Cells(sepRow + 1, startCol + 2).Value = "ACT"
    ws.Range(ws.Cells(sepRow + 1, startCol + 1), ws.Cells(sepRow + 1, startCol + 2)).Interior.Color = headerGrey
    ws.Range(ws.Cells(sepRow + 1, startCol + 1), ws.Cells(sepRow + 1, startCol + 2)).Font.Bold = True
    
    ws.Cells(sepRow + 2, startCol + 1).Value = IIf(n > 0, sumWT / n, 0)
    ws.Cells(sepRow + 2, startCol + 2).Value = IIf(n > 0, sumCT / n, 0)
    
    ws.Cells(sepRow + 2, startCol + 1).NumberFormat = "0.00"
    ws.Cells(sepRow + 2, startCol + 2).NumberFormat = "0.00"
    
    Dim areaAvg As Range
    Set areaAvg = ws.Range(ws.Cells(sepRow + 1, startCol + 1), ws.Cells(sepRow + 2, startCol + 2))
    With areaAvg.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    areaAvg.HorizontalAlignment = xlCenter
    areaAvg.VerticalAlignment = xlCenter
    
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(sepRow + 2, startCol + 2)).Columns.AutoFit
End Sub
