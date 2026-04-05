Option Explicit

Private Const HEADER_ROW_LABELS As Long = 3   ' "일자" 등이 있는 행
Private Const HEADER_ROW_TENOR   As Long = 2   ' 만기/지표명 행
Private Const DATA_START_ROW     As Long = 4   ' 데이터 시작 행
Private Const LOAD_TIMEOUT_SECONDS As Long = 300 ' 최대 로딩 대기(초)
Private Const LOAD_MIN_WAIT_SECONDS As Long = 20 ' 너무 빨리 저장되는 것 방지(초)
Private Const HEADER_SEARCH_MAX_ROWS As Long = 20 ' 새 데이터가 들어오면 헤더 행이 바뀔 수 있음
Private Const TENOR_LABEL_SEARCH_BACK_ROWS As Long = 5 ' tenor 라벨이 상단으로 밀린 경우 대응
Private Const DATA_START_SEARCH_MAX_ROWS As Long = 80 ' 데이터 시작 행 탐색 상한

' 대시보드용 CSV 저장 위치 (통합문서 폴더가 아닌 프로젝트 폴더에 고정)
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

    On Error GoTo ErrHandle

    Set ws = ThisWorkbook.Sheets(1)

    ' 헤더 행을 자동으로 찾는다(기본값: 3행이지만, 추가 데이터로 인해 바뀔 수 있음)
    Dim headerRowDate As Long
    Dim headerRowTenor As Long
    headerRowDate = FindHeaderRowDate(ws, HEADER_ROW_LABELS, HEADER_SEARCH_MAX_ROWS)
    headerRowTenor = headerRowDate - 1
    If headerRowTenor < 1 Then headerRowTenor = HEADER_ROW_TENOR

    lastCol = MaxLastColAcrossRows(ws, headerRowDate, HEADER_SEARCH_MAX_ROWS)
    lastRow = MaxLastRowAcrossDateCols(ws, headerRowDate, lastCol, DATA_START_ROW)

    Dim dataStartRow As Long
    dataStartRow = headerRowDate + 1
    If lastRow < dataStartRow Then
        MsgBox "데이터 행이 없습니다.", vbExclamation
        Exit Sub
    End If

    ' date header row에서 "일자" 컬럼 인덱스 수집
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
        MsgBox "헤더 행(3행)에서 '일자' 컬럼을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 새로 추가된 섹션으로 인해 데이터 시작 행이 달라질 수 있으므로, 실제 첫 값이 있는 행을 재탐색
    dataStartRow = FindDataStartRow(ws, dateCols, headerRowDate + 1, lastRow, DATA_START_SEARCH_MAX_ROWS)

    ' 수식/외부데이터 로딩 완료까지 대기 (최소 대기 포함)
    WaitForSheetToLoad ws, dateCols(0), dataStartRow, LOAD_TIMEOUT_SECONDS, LOAD_MIN_WAIT_SECONDS

    outPath = MARKET_DB_CSV_PATH
    numRows = 0

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText ChrW(65279)
    stream.WriteText "date,instrument,tenor,yield" & vbCrLf

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
        instrument = GetBlockInstrument(ws, headerRowTenor, dateCol, blockEndCol, prevDateCol, i + 1)

        ' tenor 라벨은 헤더에서 미리 읽어두고, export 중에는 반복 조회를 줄임
        Dim tenorStartCol As Long, tenorEndCol As Long
        Dim tenorLabels() As String
        Dim tenorCount As Long
        tenorStartCol = dateCol + 1
        tenorEndCol = blockEndCol
        tenorCount = tenorEndCol - tenorStartCol + 1

        If tenorCount > 0 Then
            ReDim tenorLabels(0 To tenorCount - 1)
            For c = tenorStartCol To tenorEndCol
                tenorLabels(c - tenorStartCol) = GetTenorLabel(ws, c, headerRowTenor, TENOR_LABEL_SEARCH_BACK_ROWS)
            Next c
        End If

        For r = dataStartRow To lastRow
            dateVal = Trim(CStr(ws.Cells(r, dateCol).Value))
            If Len(dateVal) = 0 Then GoTo NextRowBlock

            For c = tenorStartCol To tenorEndCol
                tenor = vbNullString
                If tenorCount > 0 Then tenor = tenorLabels(c - tenorStartCol)
                If Len(tenor) > 0 Then
                    cellVal = ws.Cells(r, c).Value
                    If Not (IsEmpty(cellVal) Or IsNull(cellVal) Or IsError(cellVal)) Then
                        line = EscapeCsv(dateVal) & "," & EscapeCsv(instrument) & "," & EscapeCsv(tenor) & "," & EscapeCsv(CStr(cellVal))
                        stream.WriteText line & vbCrLf
                        numRows = numRows + 1
                    End If
                End If

            Next c
NextRowBlock:
        Next r
    Next i

    stream.SaveToFile outPath, 2
    stream.Close
    MsgBox "저장 완료: " & outPath & vbCrLf & "행 수: " & numRows, vbInformation
    Exit Sub

ErrHandle:
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

Private Function QualifyPath(ByVal p As String) As String
    If Len(p) = 0 Then QualifyPath = vbNullString: Exit Function
    If Right(p, 1) = "\" Then QualifyPath = p: Exit Function
    QualifyPath = p & "\"
End Function

Private Function EscapeCsv(ByVal s As String) As String
    Dim t As String
    t = Trim(s)
    If InStr(1, t, ",") > 0 Or InStr(1, t, """") > 0 Or InStr(1, t, vbLf) > 0 Or InStr(1, t, vbCr) > 0 Then
        EscapeCsv = """" & Replace(t, """", """""") & """"
    Else
        EscapeCsv = t
    End If
End Function

Private Function GetBlockInstrument(ByVal ws As Worksheet, ByVal headerRowTenor As Long, ByVal dateCol As Long, ByVal blockEndCol As Long, ByVal prevDateCol As Long, ByVal blockNo As Long) As String
    Dim v As Variant
    Dim s As String
    Dim c As Long
    Dim rr As Long

    ' First: try to find the merged header value by scanning LEFT from dateCol
    ' (when headers are merged, only the leftmost cell holds the value).
    For rr = headerRowTenor To headerRowTenor - TENOR_LABEL_SEARCH_BACK_ROWS Step -1
        If rr < 1 Then Exit For
        For c = dateCol To prevDateCol + 1 Step -1
            v = ws.Cells(rr, c).Value
            If Not IsError(v) Then
                s = Trim(CStr(v))
                If Len(s) > 0 Then
                    If Not IsNumeric(s) Then
                        GetBlockInstrument = s
                        Exit Function
                    End If
                End If
            End If
        Next c
    Next rr

    ' Fallback: scan within the block for the first non-empty, non-error label
    For rr = headerRowTenor To headerRowTenor - TENOR_LABEL_SEARCH_BACK_ROWS Step -1
        If rr < 1 Then Exit For
        For c = dateCol To blockEndCol
            v = ws.Cells(rr, c).Value
            If Not IsError(v) Then
                s = Trim(CStr(v))
                If Len(s) > 0 Then
                    If Not IsNumeric(s) Then
                        GetBlockInstrument = s
                        Exit Function
                    End If
                End If
            End If
        Next c
    Next rr

    GetBlockInstrument = "Block" & blockNo
End Function

Private Function GetTenorLabel(ByVal ws As Worksheet, ByVal tenorCol As Long, ByVal headerRowTenor As Long, ByVal backRows As Long) As String
    Dim v As Variant
    Dim s As String
    Dim rr As Long

    ' 1) 우선 tenor 헤더 행(headerRowTenor)에서 바로 찾기
    v = ws.Cells(headerRowTenor, tenorCol).Value
    If Not IsError(v) Then
        s = Trim$(CStr(v))
        If Len(s) > 0 And InStr(1, s, "일자", vbTextCompare) = 0 And Not LooksLikeIsoDate(s) Then
            GetTenorLabel = s
            Exit Function
        End If
    End If

    ' 2) tenor 라벨이 위로 밀렸을 수도 있으니 위쪽 몇 행을 탐색
    For rr = headerRowTenor - 1 To headerRowTenor - backRows Step -1
        If rr < 1 Then Exit For
        v = ws.Cells(rr, tenorCol).Value
        If Not IsError(v) Then
            s = Trim$(CStr(v))
            If Len(s) > 0 And InStr(1, s, "일자", vbTextCompare) = 0 And Not LooksLikeIsoDate(s) Then
                GetTenorLabel = s
                Exit Function
            End If
        End If
    Next rr

    GetTenorLabel = vbNullString
End Function

Private Function LooksLikeIsoDate(ByVal s As String) As Boolean
    ' matches basic "YYYY-MM-DD" (digits + hyphens)
    Dim t As String
    t = Trim$(s)
    If Len(t) <> 10 Then
        LooksLikeIsoDate = False
        Exit Function
    End If
    If Mid$(t, 5, 1) <> "-" Then
        LooksLikeIsoDate = False
        Exit Function
    End If
    If Mid$(t, 8, 1) <> "-" Then
        LooksLikeIsoDate = False
        Exit Function
    End If
    If Not IsNumeric(Left$(t, 4)) Then
        LooksLikeIsoDate = False
        Exit Function
    End If
    If Not IsNumeric(Mid$(t, 6, 2)) Then
        LooksLikeIsoDate = False
        Exit Function
    End If
    If Not IsNumeric(Right$(t, 2)) Then
        LooksLikeIsoDate = False
        Exit Function
    End If
    LooksLikeIsoDate = True
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

    lastCol = ws.Cells(defaultRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 200

    bestRow = defaultRow
    bestCount = -1

    For rr = 1 To maxRows
        cnt = 0
        For cc = 1 To lastCol
            v = ws.Cells(rr, cc).Value
            If Not IsError(v) Then
                s = Trim$(CStr(v))
                If Len(s) > 0 Then
                    If InStr(1, s, "일자", vbTextCompare) > 0 Then
                        cnt = cnt + 1
                    End If
                End If
            End If
        Next cc
        If cnt > bestCount Then
            bestCount = cnt
            bestRow = rr
        End If
    Next rr

    FindHeaderRowDate = bestRow
End Function

Private Function MaxLastColAcrossRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal maxRows As Long) As Long
    Dim rr As Long, tmp As Long, best As Long
    best = 0
    For rr = startRow - 10 To startRow + maxRows
        If rr < 1 Then GoTo ContinueLoop
        tmp = ws.Cells(rr, ws.Columns.Count).End(xlToLeft).Column
        If tmp > best Then best = tmp
ContinueLoop:
    Next rr
    If best < 1 Then best = ws.Cells(startRow, ws.Columns.Count).End(xlToLeft).Column
    MaxLastColAcrossRows = best
End Function

Private Function MaxLastRowAcrossDateCols(ByVal ws As Worksheet, ByVal headerRowDate As Long, ByVal lastCol As Long, ByVal fallbackStartRow As Long) As Long
    Dim cc As Long, v As Variant
    Dim best As Long
    Dim tmpRow As Long
    best = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If best < fallbackStartRow Then best = fallbackStartRow

    For cc = 1 To lastCol
        v = ws.Cells(headerRowDate, cc).Value
        If Not IsError(v) Then
            If InStr(1, Trim$(CStr(v)), "일자", vbTextCompare) > 0 Then
                tmpRow = ws.Cells(ws.Rows.Count, cc).End(xlUp).Row
                If tmpRow > best Then best = tmpRow
            End If
        End If
    Next cc

    MaxLastRowAcrossDateCols = best
End Function

Private Sub WaitForSheetToLoad(ByVal ws As Worksheet, ByVal firstDateCol As Long, ByVal dataStartRow As Long, ByVal timeoutSeconds As Long, ByVal minWaitSeconds As Long)
    Dim t0 As Single, lastChange As Single
    Dim prevSig As String, sig As String
    Dim startedAsyncWait As Boolean

    t0 = Timer
    lastChange = Timer
    prevSig = ""
    startedAsyncWait = False

    On Error Resume Next

    ' Ensure calc is automatic for reliable CalculationState behavior
    Application.Calculation = xlCalculationAutomatic

    ' Trigger refresh/calc
    ThisWorkbook.RefreshAll
    Application.CalculateFullRebuild

    ' If available (newer Excel), this blocks until async queries complete
    Application.CalculateUntilAsyncQueriesDone

    Do
        DoEvents

        ' Build a lightweight signature from a few key cells that should be filled when loaded.
        ' (date cell + first value cell in the first block)
        sig = CStr(ws.Cells(dataStartRow, firstDateCol).Text) & "|" & _
              CStr(ws.Cells(dataStartRow, firstDateCol + 1).Text) & "|" & _
              CStr(ws.Cells(dataStartRow + 1, firstDateCol + 1).Text)

        If sig <> prevSig Then
            prevSig = sig
            lastChange = Timer
        End If

        ' Wait at least minWaitSeconds, and also wait until:
        ' - calculation done
        ' - connections not refreshing
        ' - and the signature has been stable for a short period (3s)
        If (Timer - t0) >= minWaitSeconds Then
            If Application.CalculationState = xlDone And Not AnyConnectionsRefreshing() Then
                If (Timer - lastChange) >= 3 Then
                    ' Also ensure the key cells are non-empty (loaded)
                    If Len(Trim$(ws.Cells(dataStartRow, firstDateCol).Text)) > 0 And _
                       Len(Trim$(ws.Cells(dataStartRow, firstDateCol + 1).Text)) > 0 Then
                        Exit Do
                    End If
                End If
            End If
        End If

        If (Timer - t0) >= timeoutSeconds Then Exit Do
    Loop

    On Error GoTo 0
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
