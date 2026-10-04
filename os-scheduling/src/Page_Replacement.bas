Option Explicit

' ---------------- CONFIG ----------------
Private Const START_ROW_REF As Long = 1
Private Const START_COL_REF As Long = 1

Private Const LEFT_LABEL_WIDTH As Double = 2.5   ' Columna B
Private Const FRAME_LABEL_WIDTH As Double = 22   ' Columna C
Private Const COL_W_NRU As Double = 6            ' NRU
Private Const COL_W_SIMPLE As Double = 3.5       ' LRU/FIFO

' ---------------- HELPERS ----------------
Private Function TrimStr(s As String) As String
    TrimStr = Trim$(s)
End Function

Private Function ParseRef(ByVal raw As Variant, ByRef page As Long, ByRef isWrite As Boolean) As Boolean
    Dim s As String
    isWrite = False
    ParseRef = False
    If IsError(raw) Then Exit Function
    If IsNull(raw) Then Exit Function
    If TrimStr(CStr(raw)) = "" Then Exit Function
    
    s = TrimStr(CStr(raw))
    If IsNumeric(s) Then
        page = CLng(s)
        ParseRef = True
        Exit Function
    End If
    
    If InStr(1, s, "-") > 0 Then
        Dim parts() As String
        parts = Split(s, "-")
        On Error GoTo ErrParse
        If IsNumeric(parts(0)) Then
            page = CLng(TrimStr(parts(0)))
            If UBound(parts) >= 1 Then
                Dim flags As String: flags = UCase$(parts(1))
                If InStr(flags, "W") > 0 Or InStr(flags, "M") > 0 Then isWrite = True
            End If
            ParseRef = True
            Exit Function
        End If
    End If
ErrParse:
    ParseRef = False
End Function

Private Function IndexOfPage(ByVal p As Long, ByRef pagesArr() As Long, ByVal framesCount As Long) As Long
    Dim idx As Long
    For idx = 1 To framesCount
        If pagesArr(idx) = p Then
            IndexOfPage = idx
            Exit Function
        End If
    Next idx
    IndexOfPage = 0
End Function

' ---------------- SIMULADORES ----------------

Public Sub SimulateNRU(ByRef refs() As Variant, ByVal nRefs As Long, ByVal framesCount As Long, _
                       ByRef out_pages() As Long, ByRef out_R() As Byte, ByRef out_M() As Byte, ByRef out_faults() As Boolean)
    Dim i As Long, tIdx As Long
    Dim frames_page() As Long, frames_R() As Byte, frames_M() As Byte
    Dim frames_loaded_at() As Long
    ReDim frames_page(1 To framesCount)
    ReDim frames_R(1 To framesCount)
    ReDim frames_M(1 To framesCount)
    ReDim frames_loaded_at(1 To framesCount)
    For i = 1 To framesCount
        frames_page(i) = -1: frames_R(i) = 0: frames_M(i) = 0: frames_loaded_at(i) = -1
    Next i
    ReDim out_pages(1 To nRefs, 1 To framesCount)
    ReDim out_R(1 To nRefs, 1 To framesCount)
    ReDim out_M(1 To nRefs, 1 To framesCount)
    ReDim out_faults(1 To nRefs)
    Dim pageVal As Long, isW As Boolean
    For tIdx = 1 To nRefs
        If Not ParseRef(refs(tIdx), pageVal, isW) Then pageVal = -1
        If pageVal <> -1 Then
            Dim idxFound As Long: idxFound = IndexOfPage(pageVal, frames_page, framesCount)
            If idxFound > 0 Then
                frames_R(idxFound) = 1
                If isW Then frames_M(idxFound) = 1
                out_faults(tIdx) = False
            Else
                out_faults(tIdx) = True
                Dim freeIdx As Long: freeIdx = 0
                For i = 1 To framesCount
                    If frames_page(i) = -1 Then: freeIdx = i: Exit For
                Next i
                If freeIdx = 0 Then
                    Dim bestCls As Long: bestCls = 999
                    Dim cands() As Long: ReDim cands(0)
                    Dim candCount As Long: candCount = 0
                    For i = 1 To framesCount
                        Dim cls As Long: cls = 2 * frames_R(i) + frames_M(i)
                        If cls < bestCls Then
                            bestCls = cls: candCount = 1: ReDim cands(1 To 1): cands(1) = i
                        ElseIf cls = bestCls Then
                            candCount = candCount + 1: ReDim Preserve cands(1 To candCount): cands(candCount) = i
                        End If
                    Next i
                    Dim victim As Long: victim = cands(1)
                    Dim bestLoaded As Long: bestLoaded = frames_loaded_at(victim)
                    Dim j As Long
                    For j = 2 To candCount
                        Dim cand As Long: cand = cands(j)
                        If bestLoaded = -1 Then
                            bestLoaded = frames_loaded_at(cand): victim = cand
                        ElseIf frames_loaded_at(cand) <> -1 And frames_loaded_at(cand) < bestLoaded Then
                            bestLoaded = frames_loaded_at(cand): victim = cand
                        End If
                    Next j
                    freeIdx = victim
                    frames_R(freeIdx) = 0: frames_M(freeIdx) = 0
                End If
                frames_page(freeIdx) = pageVal
                frames_R(freeIdx) = 1
                frames_M(freeIdx) = IIf(isW, 1, 0)
                frames_loaded_at(freeIdx) = tIdx - 1
            End If
        Else
            out_faults(tIdx) = False
        End If
        For i = 1 To framesCount
            out_pages(tIdx, i) = frames_page(i)
            out_R(tIdx, i) = frames_R(i)
            out_M(tIdx, i) = frames_M(i)
        Next i
    Next tIdx
End Sub

Public Sub SimulateLRUExact(ByRef refs() As Variant, ByVal nRefs As Long, ByVal framesCount As Long, _
                            ByRef out_pages() As Long, ByRef out_faults() As Boolean)
    Dim i As Long, tIdx As Long
    Dim frames_page() As Long, frames_last_used() As Long
    ReDim frames_page(1 To framesCount): ReDim frames_last_used(1 To framesCount)
    For i = 1 To framesCount: frames_page(i) = -1: frames_last_used(i) = -1: Next i
    ReDim out_pages(1 To nRefs, 1 To framesCount): ReDim out_faults(1 To nRefs)
    Dim pageVal As Long, isW As Boolean
    For tIdx = 1 To nRefs
        If Not ParseRef(refs(tIdx), pageVal, isW) Then pageVal = -1
        If pageVal <> -1 Then
            Dim idxFound As Long: idxFound = IndexOfPage(pageVal, frames_page, framesCount)
            If idxFound > 0 Then
                frames_last_used(idxFound) = tIdx: out_faults(tIdx) = False
            Else
                out_faults(tIdx) = True
                Dim freeIdx As Long: freeIdx = 0
                For i = 1 To framesCount
                    If frames_page(i) = -1 Then: freeIdx = i: Exit For
                Next i
                If freeIdx = 0 Then
                    Dim victim As Long: victim = 1: Dim bestLU As Long: bestLU = frames_last_used(1)
                    For i = 2 To framesCount
                        If frames_last_used(i) < bestLU Then: bestLU = frames_last_used(i): victim = i
                    Next i
                    freeIdx = victim
                End If
                frames_page(freeIdx) = pageVal: frames_last_used(freeIdx) = tIdx
            End If
        End If
        For i = 1 To framesCount: out_pages(tIdx, i) = frames_page(i): Next i
    Next tIdx
End Sub

Public Sub SimulateFIFO_Requeue(ByRef refs() As Variant, ByVal nRefs As Long, ByVal framesCount As Long, _
                                ByRef out_pages() As Long, ByRef out_faults() As Boolean)
    Dim i As Long, tIdx As Long
    Dim frames_page() As Long
    ReDim frames_page(1 To framesCount)
    For i = 1 To framesCount: frames_page(i) = -1: Next i
    ReDim out_pages(1 To nRefs, 1 To framesCount): ReDim out_faults(1 To nRefs)
    
    Dim queue() As Long
    ReDim queue(1 To framesCount)
    Dim qCount As Long: qCount = 0
    
    Dim pageVal As Long, isW As Boolean
    
    For tIdx = 1 To nRefs
        If Not ParseRef(refs(tIdx), pageVal, isW) Then pageVal = -1
        
        If pageVal <> -1 Then
            Dim idxFound As Long: idxFound = IndexOfPage(pageVal, frames_page, framesCount)
            
            If idxFound > 0 Then
                out_faults(tIdx) = False
                
                Dim qPos As Long, k As Long
                Dim foundInQ As Boolean: foundInQ = False
                For k = 1 To qCount
                    If queue(k) = idxFound Then
                        qPos = k
                        foundInQ = True
                        Exit For
                    End If
                Next k
                
                If foundInQ Then
                    For k = qPos To qCount - 1
                        queue(k) = queue(k + 1)
                    Next k
                    qCount = qCount - 1
                End If
                qCount = qCount + 1
                queue(qCount) = idxFound
                
            Else
                out_faults(tIdx) = True
                Dim targetFrame As Long
                
                If qCount < framesCount Then
                    For i = 1 To framesCount
                        If frames_page(i) = -1 Then
                            targetFrame = i
                            Exit For
                        End If
                    Next i
                Else
                    targetFrame = queue(1)
                    Dim m As Long
                    For m = 1 To qCount - 1
                        queue(m) = queue(m + 1)
                    Next m
                    qCount = qCount - 1
                End If
                
                frames_page(targetFrame) = pageVal
                qCount = qCount + 1
                queue(qCount) = targetFrame
            End If
        Else
             out_faults(tIdx) = False
        End If
        
        For i = 1 To framesCount: out_pages(tIdx, i) = frames_page(i): Next i
    Next tIdx
End Sub

' ---------------- FORMAT AND WRITE (VISUALS) ----------------
Public Sub FormatAndWriteTablesToSheet( _
    ByVal sh As Worksheet, _
    ByRef refs As Variant, _
    ByVal nRefs As Long, _
    ByVal framesCount As Long, _
    ByVal algo As String, _
    ByRef pagesNRU As Variant, _
    ByRef Rs As Variant, _
    ByRef Ms As Variant, _
    ByRef faultsNRU As Variant, _
    ByRef pagesSimple As Variant, _
    ByRef faultsSimple As Variant, _
    ByVal outCell As String)

    Dim rngStart As Range
    Set rngStart = sh.Range(outCell)

    Dim startRow As Long: startRow = rngStart.Row
    Dim startCol As Long: startCol = rngStart.Column
    
    Dim dataColStart As Long: dataColStart = startCol + 2
    Dim lastCol As Long: lastCol = dataColStart + nRefs - 1
    
    ' --- ESTADISTICAS ---
    Dim totalFaults As Long: totalFaults = 0
    Dim tIdx As Long
    
    If UCase$(algo) = "NRU" Then
        For tIdx = 1 To nRefs
            If faultsNRU(tIdx) Then totalFaults = totalFaults + 1
        Next tIdx
    Else
        For tIdx = 1 To nRefs
            If faultsSimple(tIdx) Then totalFaults = totalFaults + 1
        Next tIdx
    End If
    
    Dim tasaFallo As Double
    If nRefs > 0 Then tasaFallo = totalFaults / nRefs Else tasaFallo = 0
    Dim rendimiento As Double: rendimiento = 1 - tasaFallo

    ' --- VISUALIZACION ---
    Dim rowsPerFrame As Long
    Dim spanCols As Integer
    
    If UCase$(algo) = "NRU" Then
        rowsPerFrame = 3
        spanCols = 2
    Else
        rowsPerFrame = 2
        spanCols = 3
    End If
    
    Dim framesBlockRows As Long
    framesBlockRows = framesCount * rowsPerFrame

    Application.ScreenUpdating = False

    Dim coloBlueHeader As Long: coloBlueHeader = RGB(218, 233, 248)
    Dim coloGrey As Long: coloGrey = RGB(242, 242, 242)
    Dim coloFailBg As Long: coloFailBg = RGB(251, 226, 213)
    Dim coloRedText As Long: coloRedText = RGB(255, 0, 0)
    Dim coloWhite As Long: coloWhite = RGB(255, 255, 255)

    sh.Range(sh.Cells(startRow, startCol), sh.Cells(startRow + 400, lastCol + 10)).Clear
    With sh.Cells.Font
        .Name = "Calibri"
        .Size = 11
        .Color = vbBlack
    End With

    sh.Columns(startCol).ColumnWidth = LEFT_LABEL_WIDTH
    sh.Columns(startCol + 1).ColumnWidth = FRAME_LABEL_WIDTH
    
    Dim c As Long
    Dim currentDataWidth As Double
    If UCase$(algo) = "NRU" Then currentDataWidth = COL_W_NRU Else currentDataWidth = COL_W_SIMPLE
    
    For c = dataColStart To lastCol
        sh.Columns(c).ColumnWidth = currentDataWidth
    Next c

    ' ---------------- TABLA SUPERIOR (INFO) ----------------
    With sh.Range(sh.Cells(startRow, startCol), sh.Cells(startRow, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = algo
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = coloBlueHeader
        .Borders.LineStyle = xlContinuous
    End With

    With sh.Range(sh.Cells(startRow + 1, startCol), sh.Cells(startRow + 1, startCol + spanCols - 1))
        .Merge
        .Value = "Número de Referencias:"
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
    End With
    With sh.Range(sh.Cells(startRow + 1, startCol + spanCols), sh.Cells(startRow + 1, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = nRefs
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
    End With

    With sh.Range(sh.Cells(startRow + 2, startCol), sh.Cells(startRow + 2, startCol + spanCols - 1))
        .Merge
        .Value = "Número de Frames:"
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
    End With
    With sh.Range(sh.Cells(startRow + 2, startCol + spanCols), sh.Cells(startRow + 2, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = framesCount
        .HorizontalAlignment = xlCenter
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
    End With

    Dim gridStartRow As Long
    gridStartRow = startRow + 5

    ' ---------------- HEADER: Memoria virtual ----------------
    With sh.Range(sh.Cells(gridStartRow, dataColStart), sh.Cells(gridStartRow, lastCol))
        .Merge
        .Value = "Memoria virtual"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    ' ---------------- ROW: Referencias (Flujo) ----------------
    Dim rowRef As Long: rowRef = gridStartRow + 1
    
    With sh.Cells(rowRef, startCol + 1)
        .Value = "Flujo de páginas"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = coloGrey
    End With
    
    For tIdx = 1 To nRefs
        With sh.Cells(rowRef, dataColStart + tIdx - 1)
            .Value = refs(tIdx)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = coloGrey
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
        End With
    Next tIdx

    Dim rowFramesStart As Long: rowFramesStart = rowRef + 1
    Dim rowFramesEnd As Long: rowFramesEnd = rowFramesStart + framesBlockRows - 1
    Dim rowFallos As Long: rowFallos = rowFramesEnd + 1

    ' ---------------- LEFT COLUMN: Memoria Fisica ----------------
    With sh.Range(sh.Cells(rowRef, startCol), sh.Cells(rowFallos, startCol))
        .Merge
        .Value = "Memoria física"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 90
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    ' ---------------- DATA FRAMES ----------------
    Dim currRow As Long: currRow = rowFramesStart
    Dim j As Long

    If UCase$(algo) = "NRU" Then
        ' --- NRU ---
        For j = 1 To framesCount
            With sh.Range(sh.Cells(currRow, startCol + 1), sh.Cells(currRow + 2, startCol + 1))
                .Merge
                .Value = "Marco/Frame " & j
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With

            For tIdx = 1 To nRefs
                Dim valPage As Long: valPage = pagesNRU(tIdx, j)
                Dim valR As Byte: valR = Rs(tIdx, j)
                Dim valM As Byte: valM = Ms(tIdx, j)
                
                Dim refPage As Long, refIsW As Boolean
                Dim isTarget As Boolean: isTarget = False
                If ParseRef(refs(tIdx), refPage, refIsW) Then
                    If valPage > -1 And valPage = refPage Then isTarget = True
                End If

                Dim rngBlock As Range
                Set rngBlock = sh.Range(sh.Cells(currRow, dataColStart + tIdx - 1), sh.Cells(currRow + 2, dataColStart + tIdx - 1))
                
                rngBlock.Interior.Color = coloWhite
                rngBlock.Borders.LineStyle = xlContinuous
                rngBlock.Borders(xlInsideHorizontal).LineStyle = xlNone
                rngBlock.HorizontalAlignment = xlCenter
                rngBlock.VerticalAlignment = xlCenter

                With sh.Cells(currRow, dataColStart + tIdx - 1)
                    If valPage = -1 Then .Value = "" Else .Value = valPage
                    If isTarget Then .Font.Bold = True
                    If isTarget And faultsNRU(tIdx) Then .Font.Color = coloRedText
                End With
                
                With sh.Cells(currRow + 1, dataColStart + tIdx - 1)
                    If valPage <> -1 Then
                        .Value = "R=" & valR
                        If isTarget Then .Font.Color = coloRedText
                    End If
                End With
                
                With sh.Cells(currRow + 2, dataColStart + tIdx - 1)
                    If valPage <> -1 Then
                        .Value = "M=" & valM
                        If isTarget And refIsW Then .Font.Color = coloRedText
                    End If
                End With
            Next tIdx
            currRow = currRow + 3
        Next j
        
    Else
        ' --- SIMPLES (LRU / FIFO) ---
        For j = 1 To framesCount
            With sh.Range(sh.Cells(currRow, startCol + 1), sh.Cells(currRow + 1, startCol + 1))
                .Merge
                .Value = "Marco/Frame " & j
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
            
            For tIdx = 1 To nRefs
                Dim valP As Long: valP = pagesSimple(tIdx, j)
                Dim refPageS As Long, refIsWS As Boolean
                Dim isTargetS As Boolean: isTargetS = False
                If ParseRef(refs(tIdx), refPageS, refIsWS) Then
                    If valP > -1 And valP = refPageS Then isTargetS = True
                End If
                
                Dim rngBlockS As Range
                Set rngBlockS = sh.Range(sh.Cells(currRow, dataColStart + tIdx - 1), sh.Cells(currRow + 1, dataColStart + tIdx - 1))
                
                With rngBlockS
                    .Merge
                    If valP = -1 Then .Value = "" Else .Value = valP
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Interior.Color = coloWhite
                    
                    If isTargetS Then .Font.Bold = True
                    If isTargetS And faultsSimple(tIdx) Then .Font.Color = coloRedText
                End With
            Next tIdx
            currRow = currRow + 2
        Next j
    End If

    ' ---------------- FALLOS ROW ----------------
    With sh.Cells(rowFallos, startCol + 1)
        .Value = "Fallos"
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .Interior.Color = coloFailBg
    End With
    
    With sh.Range(sh.Cells(rowFallos, dataColStart), sh.Cells(rowFallos, lastCol))
        .Interior.Color = coloFailBg
        .Borders.LineStyle = xlContinuous
    End With
    
    For tIdx = 1 To nRefs
        Dim showF As Boolean
        If UCase$(algo) = "NRU" Then showF = faultsNRU(tIdx) Else showF = faultsSimple(tIdx)
        
        With sh.Cells(rowFallos, dataColStart + tIdx - 1)
            If showF Then
                .Value = "F"
                .Font.Color = coloRedText
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            Else
                .Value = ""
            End If
        End With
    Next tIdx
    
    ' ---------------- TABLA DE RESULTADOS (ABAJO) ----------------
    Dim resRow As Long
    resRow = rowFallos + 3
    
    ' -- HEADER --
    With sh.Range(sh.Cells(resRow, startCol), sh.Cells(resRow, startCol + spanCols - 1))
        .Merge
        .Value = "Items"
        .Interior.Color = coloGrey
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    With sh.Range(sh.Cells(resRow, startCol + spanCols), sh.Cells(resRow, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = "Resultados"
        .Interior.Color = coloGrey
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' -- DATA --
    With sh.Range(sh.Cells(resRow + 1, startCol), sh.Cells(resRow + 1, startCol + spanCols - 1))
        .Merge
        .Value = "Número de Fallos:"
        .Borders.LineStyle = xlContinuous
    End With
    With sh.Range(sh.Cells(resRow + 1, startCol + spanCols), sh.Cells(resRow + 1, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = totalFaults
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    With sh.Range(sh.Cells(resRow + 2, startCol), sh.Cells(resRow + 2, startCol + spanCols - 1))
        .Merge
        .Value = "Porcentaje de Fallos:"
        .Borders.LineStyle = xlContinuous
    End With
    With sh.Range(sh.Cells(resRow + 2, startCol + spanCols), sh.Cells(resRow + 2, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = Format(tasaFallo, "0.00%")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    With sh.Range(sh.Cells(resRow + 3, startCol), sh.Cells(resRow + 3, startCol + spanCols - 1))
        .Merge
        .Value = "Rendimiento (%):"
        .Borders.LineStyle = xlContinuous
    End With
    With sh.Range(sh.Cells(resRow + 3, startCol + spanCols), sh.Cells(resRow + 3, startCol + (2 * spanCols) - 1))
        .Merge
        .Value = Format(rendimiento, "0.00%")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    Application.ScreenUpdating = True
End Sub

' ---------------- RUN ----------------
Public Sub RunPagingSimulation()
    Dim sh As Worksheet
    Set sh = ActiveSheet

    Dim refsArr() As Variant
    Dim colIdx As Long: colIdx = START_COL_REF
    Dim nRefs As Long: nRefs = 0
    
    Do
        Dim v As Variant
        v = sh.Cells(START_ROW_REF, colIdx).Value
        If TrimStr(CStr(v)) = "" Then Exit Do
        nRefs = nRefs + 1
        ReDim Preserve refsArr(1 To nRefs)
        refsArr(nRefs) = v
        colIdx = colIdx + 1
    Loop

    If nRefs = 0 Then
        MsgBox "No se encontraron referencias en fila " & START_ROW_REF, vbExclamation
        Exit Sub
    End If

    Dim algoChoice As String
    algoChoice = InputBox("Algoritmo (NRU / LRU / FIFO):", "Simulador", "NRU")
    If Trim$(algoChoice) = "" Then Exit Sub
    algoChoice = UCase$(Trim$(algoChoice))

    Dim framesCount As Long
    framesCount = CLng(Application.InputBox("Numero de frames:", "Simulador", 3, Type:=1))
    If framesCount <= 0 Then Exit Sub

    Dim outCell As String
    outCell = "B" & (START_ROW_REF + 4)

    Dim pagesNRU() As Long, Rs() As Byte, Ms() As Byte, faultsNRU() As Boolean
    Dim pagesSimple() As Long, faultsSimple() As Boolean

    If algoChoice = "NRU" Then
        SimulateNRU refsArr, nRefs, framesCount, pagesNRU, Rs, Ms, faultsNRU
    ElseIf algoChoice = "LRU" Then
        SimulateLRUExact refsArr, nRefs, framesCount, pagesSimple, faultsSimple
    ElseIf algoChoice = "FIFO" Then
        SimulateFIFO_Requeue refsArr, nRefs, framesCount, pagesSimple, faultsSimple
    Else
        MsgBox "Algoritmo desconocido"
        Exit Sub
    End If

    Call FormatAndWriteTablesToSheet( _
        sh, _
        refsArr, nRefs, framesCount, algoChoice, _
        pagesNRU, Rs, Ms, faultsNRU, _
        pagesSimple, faultsSimple, _
        outCell)

    MsgBox "Completado.", vbInformation
End Sub

