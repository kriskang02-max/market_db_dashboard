Option Explicit

Private Const HEADER_ROW_LABELS As Long = 3
Private Const HEADER_ROW_TENOR   As Long = 2
Private Const DATA_START_ROW     As Long = 4
Private Const LOAD_TIMEOUT_SECONDS As Long = 45
Private Const LOAD_MIN_WAIT_SECONDS As Long = 3
Private Const HEADER_SEARCH_MAX_ROWS As Long = 20
Private Const TENOR_LABEL_SEARCH_BACK_ROWS As Long = 5
Private Const DATA_START_SEARCH_MAX_ROWS As Long = 80
Private Const EXPORT_FLUSH_BATCH As Long = 25000
Private Const SKIP_WAIT_IF_LOADED As Boolean = True
Private Const FORCE_REFRESH_BEFORE_EXPORT As Boolean = False

Private Const MARKET_DB_CSV_PATH As String = "C:\Users\infomax\Documents\market_db_dashboard\market_db.csv"

Public Sub ExportMarketDbLong()
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim dateCols() As Long
    Dim nBlocks As Long
    Dim i As Long, j As Long, r As Long, c As Long
    Dim dateCol As Long, blockEndCol As Long
    Dim instrument As String, tenor As String, dateVal As String, cellVal As Variant
    Dim outPath As String
    Dim line As String
    Dim stream As Object
    Dim numRows As Long
    Dim allDat As Variant
    Dim hdrDat As Variant
    Dim batchLines() As String
    Dim batchCount As Long
    Dim savedScreenUpdating As Boolean
    Dim savedEnableEvents As Boolean
    Dim savedCalculation As XlCalculation
    Dim savedStatusBar As Boolean
    Dim escapedInst As String
    Dim escapedTenors() As String
    Dim tenorStartCol As Long, tenorEndCol As Long
    Dim tenorCount As Long
    Dim hdrRows As Long

    On Error GoTo ErrHandle

    savedScreenUpdating = Application.ScreenUpdating
    savedEnableEvents = Application.EnableEvents
    savedCalculation = Application.Calculation
    savedStatusBar = Application.DisplayStatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True

    Set ws = ThisWorkbook.Sheets(1)

    SetStatus "CSV보내기: 헤더 확인 중..."
    DoEvents

    Dim headerRowDate As Long
    Dim headerRowTenor As Long
    headerRowDate = FindHeaderRowDate(ws, HEADER_ROW_LABELS, HEADER_SEARCH_MAX_ROWS)
    headerRowTenor = headerRowDate - 1
    If headerRowTenor < 1 Then headerRowTenor = HEADER_ROW_TENOR

    lastCol = MaxLastColAcrossRows(ws, headerRowDate, HEADER_SEARCH_MAX_ROWS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW

    Dim dataStartRow As Long
    dataStartRow = headerRowDate + 1
    If lastRow < dataStartRow Then
        MsgBox "데이터 행이 없습니다.", vbExclamation
        GoTo CleanExit
    End If

    ReDim dateCols(0 To 0)
    nBlocks = 0
    For j = 1 To lastCol
        If InStr(1, Trim$(CStr(ws.Cells(headerRowDate, j).Value)), "일자", vbTextCompare) > 0 Then
            If nBlocks > 0 Then ReDim Preserve dateCols(0 To nBlocks)
            dateCols(nBlocks) = j
            nBlocks = nBlocks + 1
        End If
    Next j

    If nBlocks = 0 Then
        MsgBox "헤더 행에서 '일자' 컬럼을 찾을 수 없습니다.", vbExclamation
        GoTo CleanExit
    End If

    lastRow = MaxLastRowAcrossDateCols(ws, dateCols, DATA_START_ROW)
    dataStartRow = FindDataStartRow(ws, dateCols, headerRowDate + 1, lastRow, DATA_START_SEARCH_MAX_ROWS)

    DisableBackgroundRefresh
    WaitForSheetToLoad ws, dateCols(0), dataStartRow, LOAD_TIMEOUT_SECONDS, LOAD_MIN_WAIT_SECONDS

    lastRow = MaxLastRowAcrossDateCols(ws, dateCols, DATA_START_ROW)
    If lastRow < dataStartRow Then
        MsgBox "로딩 후 데이터 행이 없습니다.", vbExclamation
        GoTo CleanExit
    End If

    SetStatus "CSV보내기: 시트 데이터 읽는 중..."
    DoEvents

    hdrRows = headerRowTenor + TENOR_LABEL_SEARCH_BACK_ROWS
    If hdrRows < headerRowDate Then hdrRows = headerRowDate
    hdrDat = ws.Range(ws.Cells(1, 1), ws.Cells(hdrRows, lastCol)).Value2
    allDat = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(lastRow, lastCol)).Value2

    outPath = MARKET_DB_CSV_PATH
    numRows = 0
    batchCount = 0
    ReDim batchLines(0 To EXPORT_FLUSH_BATCH - 1)

    SetStatus "CSV보내기: 파일 준비 중..."
    DoEvents

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText ChrW(65279), 0
    stream.WriteText "date,instrument,tenor,yield" & vbCrLf, 0

    For i = 0 To nBlocks - 1
        dateCol = dateCols(i)
        If i < nBlocks - 1 Then
            blockEndCol = dateCols(i + 1) - 1
        Else
            blockEndCol = lastCol
        End If

        Dim prevDateCol As Long
        prevDateCol = 0
        If i > 0 Then prevDateCol = dateCols(i - 1)
        instrument = GetBlockInstrumentFromHdr(hdrDat, headerRowTenor, dateCol, blockEndCol, prevDateCol, i + 1)
        escapedInst = EscapeCsv(instrument)

        tenorStartCol = dateCol + 1
        tenorEndCol = blockEndCol
        tenorCount = tenorEndCol - tenorStartCol + 1

        If tenorCount > 0 Then
            ReDim escapedTenors(0 To tenorCount - 1)
            For c = tenorStartCol To tenorEndCol
                tenor = GetTenorLabelFromHdr(hdrDat, c, headerRowTenor, TENOR_LABEL_SEARCH_BACK_ROWS)
                escapedTenors(c - tenorStartCol) = EscapeCsv(tenor)
            Next c
        End If

        For r = 1 To UBound(allDat, 1)
            dateVal = ExportFormatDateForCsv(allDat(r, dateCol))
            If Len(dateVal) = 0 Then GoTo NextRowBlock

            For c = tenorStartCol To tenorEndCol
                If Len(escapedTenors(c - tenorStartCol)) = 0 Then GoTo NextColBlock
                cellVal = allDat(r, c)
                If Not (IsEmpty(cellVal) Or IsNull(cellVal) Or IsError(cellVal)) Then
                    line = dateVal & "," & escapedInst & "," & escapedTenors(c - tenorStartCol) & "," & EscapeCsv(CStr(cellVal))
                    batchLines(batchCount) = line
                    batchCount = batchCount + 1
                    numRows = numRows + 1
                    If batchCount >= EXPORT_FLUSH_BATCH Then
                        FlushBatchToStream stream, batchLines, batchCount
                        batchCount = 0
                        SetStatus "CSV보내기: " & Format(numRows, "#,##0") & "행 처리됨..."
                        DoEvents
                    End If
                End If
NextColBlock:
            Next c
NextRowBlock:
        Next r
    Next i

    If batchCount > 0 Then FlushBatchToStream stream, batchLines, batchCount

    SetStatus "CSV보내기: 파일 저장 중..."
    DoEvents

    On Error Resume Next
    Kill outPath
    On Error GoTo ErrHandle

    stream.SaveToFile outPath, 2
    stream.Close
    Set stream = Nothing

    MsgBox "저장 완료: " & outPath & vbCrLf & "행 수: " & Format(numRows, "#,##0"), vbInformation
    GoTo CleanExit

ErrHandle:
    MsgBox "오류: " & Err.Number & " — " & Err.Description, vbCritical

CleanExit:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    Application.StatusBar = False
    Application.DisplayStatusBar = savedStatusBar
    Application.Calculation = savedCalculation
    Application.ScreenUpdating = savedScreenUpdating
    Application.EnableEvents = savedEnableEvents
    On Error GoTo 0
End Sub

Private Sub SetStatus(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Private Sub FlushBatchToStream(ByVal stream As Object, ByRef lines() As String, ByVal n As Long)
    Dim i As Long
    If n <= 0 Then Exit Sub
    If n = UBound(lines) - LBound(lines) + 1 Then
        stream.WriteText Join(lines, vbCrLf) & vbCrLf, 0
        Exit Sub
    End If
    Dim chunkArr() As String
    ReDim chunkArr(0 To n - 1)
    For i = 0 To n - 1
        chunkArr(i) = lines(i)
    Next i
    stream.WriteText Join(chunkArr, vbCrLf) & vbCrLf, 0
End Sub

Private Function ExportFormatDateForCsv(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then
        ExportFormatDateForCsv = vbNullString
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        ExportFormatDateForCsv = vbNullString
        Exit Function
    End If

    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        If IsNumeric(Left$(s, 4)) And IsNumeric(Mid$(s, 6, 2)) And IsNumeric(Right$(s, 2)) Then
            ExportFormatDateForCsv = s
            Exit Function
        End If
    End If

    If IsNumeric(v) Then
        Dim serial As Double
        serial = CDbl(v)
        If serial >= 1 And serial < 2958466 Then
            ExportFormatDateForCsv = Format$(CDate(serial), "yyyy-mm-dd")
            Exit Function
        End If
    End If

    On Error Resume Next
    ExportFormatDateForCsv = Format$(CDate(v), "yyyy-mm-dd")
    If Err.Number <> 0 Then
        Err.Clear
        ExportFormatDateForCsv = s
    End If
    On Error GoTo 0
End Function

Private Function SheetLooksLoaded(ByVal ws As Worksheet, ByVal dateCol As Long, ByVal dataStartRow As Long) As Boolean
    Dim d As String, y As String
    d = Trim$(CStr(ws.Cells(dataStartRow, dateCol).Text))
    y = Trim$(CStr(ws.Cells(dataStartRow, dateCol + 1).Text))
    SheetLooksLoaded = (Len(d) > 0 And Len(y) > 0)
End Function

Private Sub DisableBackgroundRefresh()
    Dim conn As WorkbookConnection
    Dim qtWs As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject

    On Error Resume Next
    For Each conn In ThisWorkbook.Connections
        conn.OLEDBConnection.BackgroundQuery = False
        conn.ODBCConnection.BackgroundQuery = False
    Next conn

    For Each qtWs In ThisWorkbook.Worksheets
        For Each qt In qtWs.QueryTables
            qt.BackgroundQuery = False
        Next qt
        For Each lo In qtWs.ListObjects
            lo.QueryTable.BackgroundQuery = False
        Next lo
    Next qtWs
    On Error GoTo 0
End Sub

Private Function EscapeCsv(ByVal s As String) As String
    Dim t As String
    t = Trim(s)
    If InStr(1, t, ",") > 0 Or InStr(1, t, """") > 0 Or InStr(1, t, vbLf) > 0 Or InStr(1, t, vbCr) > 0 Then
        EscapeCsv = """" & Replace(t, """", """""") & """"
    Else
        EscapeCsv = t
    End If
End Function

Private Function GetBlockInstrumentFromHdr(ByVal hdr As Variant, ByVal headerRowTenor As Long, ByVal dateCol As Long, ByVal blockEndCol As Long, ByVal prevDateCol As Long, ByVal blockNo As Long) As String
    Dim v As Variant
    Dim s As String
    Dim c As Long
    Dim rr As Long

    For rr = headerRowTenor To headerRowTenor - TENOR_LABEL_SEARCH_BACK_ROWS Step -1
        If rr < 1 Then Exit For
        For c = dateCol To prevDateCol + 1 Step -1
            v = HdrCell(hdr, rr, c)
            If Not IsError(v) Then
                s = Trim(CStr(v))
                If Len(s) > 0 And Not IsNumeric(s) Then
                    GetBlockInstrumentFromHdr = s
                    Exit Function
                End If
            End If
        Next c
    Next rr

    For rr = headerRowTenor To headerRowTenor - TENOR_LABEL_SEARCH_BACK_ROWS Step -1
        If rr < 1 Then Exit For
        For c = dateCol To blockEndCol
            v = HdrCell(hdr, rr, c)
            If Not IsError(v) Then
                s = Trim(CStr(v))
                If Len(s) > 0 And Not IsNumeric(s) Then
                    GetBlockInstrumentFromHdr = s
                    Exit Function
                End If
            End If
        Next c
    Next rr

    GetBlockInstrumentFromHdr = "Block" & blockNo
End Function

Private Function GetTenorLabelFromHdr(ByVal hdr As Variant, ByVal tenorCol As Long, ByVal headerRowTenor As Long, ByVal backRows As Long) As String
    Dim v As Variant
    Dim s As String
    Dim rr As Long

    v = HdrCell(hdr, headerRowTenor, tenorCol)
    If Not IsError(v) Then
        s = Trim$(CStr(v))
        If Len(s) > 0 And InStr(1, s, "일자", vbTextCompare) = 0 And Not LooksLikeIsoDate(s) Then
            GetTenorLabelFromHdr = s
            Exit Function
        End If
    End If

    For rr = headerRowTenor - 1 To headerRowTenor - backRows Step -1
        If rr < 1 Then Exit For
        v = HdrCell(hdr, rr, tenorCol)
        If Not IsError(v) Then
            s = Trim$(CStr(v))
            If Len(s) > 0 And InStr(1, s, "일자", vbTextCompare) = 0 And Not LooksLikeIsoDate(s) Then
                GetTenorLabelFromHdr = s
                Exit Function
            End If
        End If
    Next rr

    GetTenorLabelFromHdr = vbNullString
End Function

Private Function HdrCell(ByVal hdr As Variant, ByVal row As Long, ByVal col As Long) As Variant
    On Error Resume Next
    HdrCell = hdr(row, col)
    On Error GoTo 0
End Function

Private Function LooksLikeIsoDate(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) <> 10 Then LooksLikeIsoDate = False: Exit Function
    If Mid$(t, 5, 1) <> "-" Or Mid$(t, 8, 1) <> "-" Then LooksLikeIsoDate = False: Exit Function
    If Not IsNumeric(Left$(t, 4)) Or Not IsNumeric(Mid$(t, 6, 2)) Or Not IsNumeric(Right$(t, 2)) Then
        LooksLikeIsoDate = False
    Else
        LooksLikeIsoDate = True
    End If
End Function

Private Function FindDataStartRow(ByVal ws As Worksheet, ByRef dateCols() As Long, ByVal startRow As Long, ByVal endRow As Long, ByVal maxScanExtra As Long) As Long
    Dim r As Long, i As Long
    Dim v As Variant
    Dim s As String
    Dim scanEnd As Long

    FindDataStartRow = startRow
    If endRow < startRow Then Exit Function

    scanEnd = endRow
    If scanEnd - startRow > maxScanExtra Then scanEnd = startRow + maxScanExtra

    For r = startRow To scanEnd
        For i = LBound(dateCols) To UBound(dateCols)
            v = ws.Cells(r, dateCols(i)).Value
            If Not IsError(v) Then
                s = Trim$(CStr(v))
                If Len(s) > 0 Then
                    FindDataStartRow = r
                    Exit Function
                End If
            End If
        Next i
    Next r
End Function

Private Function FindHeaderRowDate(ByVal ws As Worksheet, ByVal defaultRow As Long, ByVal maxRows As Long) As Long
    Dim rr As Long, cc As Long
    Dim lastCol As Long
    Dim bestRow As Long, bestCount As Long, cnt As Long
    Dim v As Variant, s As String
    Dim hdrScan As Variant

    lastCol = ws.Cells(defaultRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 200

    hdrScan = ws.Range(ws.Cells(1, 1), ws.Cells(maxRows, lastCol)).Value2

    bestRow = defaultRow
    bestCount = -1

    For rr = 1 To maxRows
        cnt = 0
        For cc = 1 To lastCol
            v = hdrScan(rr, cc)
            If Not IsError(v) Then
                s = Trim$(CStr(v))
                If Len(s) > 0 And InStr(1, s, "일자", vbTextCompare) > 0 Then cnt = cnt + 1
            End If
        Next cc
        If cnt > bestCount Then bestCount = cnt: bestRow = rr
    Next rr

    FindHeaderRowDate = bestRow
End Function

Private Function MaxLastColAcrossRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal maxRows As Long) As Long
    Dim rr As Long, tmp As Long, best As Long
    best = ws.Cells(startRow, ws.Columns.Count).End(xlToLeft).Column
    For rr = startRow - 3 To startRow + 3
        If rr >= 1 Then
            tmp = ws.Cells(rr, ws.Columns.Count).End(xlToLeft).Column
            If tmp > best Then best = tmp
        End If
    Next rr
    If best < 1 Then best = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    MaxLastColAcrossRows = best
End Function

Private Function MaxLastRowAcrossDateCols(ByVal ws As Worksheet, ByRef dateCols() As Long, ByVal fallbackStartRow As Long) As Long
    Dim i As Long, tmpRow As Long
    Dim best As Long
    best = ws.Cells(ws.Rows.Count, dateCols(0)).End(xlUp).Row
    If best < fallbackStartRow Then best = fallbackStartRow

    For i = LBound(dateCols) To UBound(dateCols)
        tmpRow = ws.Cells(ws.Rows.Count, dateCols(i)).End(xlUp).Row
        If tmpRow > best Then best = tmpRow
    Next i

    MaxLastRowAcrossDateCols = best
End Function

Private Sub WaitForSheetToLoad(ByVal ws As Worksheet, ByVal firstDateCol As Long, ByVal dataStartRow As Long, ByVal timeoutSeconds As Long, ByVal minWaitSeconds As Long)
    Dim t0 As Single, lastChange As Single
    Dim prevSig As String, sig As String

    If SKIP_WAIT_IF_LOADED And SheetLooksLoaded(ws, firstDateCol, dataStartRow) Then
        SetStatus "데이터 확인됨 — 새로고침 대기 건너뜀"
        DoEvents
        Exit Sub
    End If

    SetStatus "외부 데이터 로딩 대기 중..."
    DoEvents

    t0 = Timer
    lastChange = Timer
    prevSig = vbNullString

    On Error Resume Next
    If FORCE_REFRESH_BEFORE_EXPORT Then
        ThisWorkbook.RefreshAll
        If Err.Number <> 0 Then Err.Clear
    End If
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    Do
        DoEvents

        sig = CStr(ws.Cells(dataStartRow, firstDateCol).Text) & "|" & _
              CStr(ws.Cells(dataStartRow, firstDateCol + 1).Text)

        If sig <> prevSig Then
            prevSig = sig
            lastChange = Timer
            SetStatus "외부 데이터 로딩 중... (" & Format(Timer - t0, "0") & "초)"
        End If

        If (Timer - t0) >= minWaitSeconds Then
            If Application.CalculationState = xlDone And Not AnyConnectionsRefreshing() Then
                If SheetLooksLoaded(ws, firstDateCol, dataStartRow) Then
                    If (Timer - lastChange) >= 2 Then Exit Do
                End If
            End If
        End If

        If (Timer - t0) >= timeoutSeconds Then Exit Do
    Loop
End Sub

Private Function AnyConnectionsRefreshing() As Boolean
    Dim c As WorkbookConnection
    On Error Resume Next
    AnyConnectionsRefreshing = False
    For Each c In ThisWorkbook.Connections
        If Not c Is Nothing Then
            If c.Refreshing Then
                AnyConnectionsRefreshing = True
                Exit Function
            End If
        End If
    Next c
    On Error GoTo 0
End Function
