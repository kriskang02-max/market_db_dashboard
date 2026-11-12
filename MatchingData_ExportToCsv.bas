Attribute VB_Name = "MatchingData_ExportToCsv"
' File encoding: CP949 (Excel VBA import on Korean Windows)
'===============================================================================
' Matching_data.xlsm -> 자동화용 롱 CSV
'
' 시트 구조 (인포맥스형, 가로 반복 블록):
'   행1          : 메타 (시작/종료/Data 개수/주기) - CSV에 넣지 않음
'   헤더 바로 위 : 블록 이름 (시가평가 3사평균 산업은행, CD 3사평균 AAA, REPO)
'   헤더 행      : "일자" | 항목    다음 "일자" 열까지가 한 블록 (너비/개수 가변)
'   그 아래      : 블록마다 자신의 일자 + 값 (블록마다 행 수가 달라도 됨)
'
' 출력 (통합문서와 같은 폴더):
'   C:\Users\infomax\Documents\market_db_dashboard\matching_data.csv
'   헤더: date,instrument,field,value
'   한 행 = 날짜 1개 x 종목 1개 x 항목 1개 x 값 1개
'   날짜는 yyyy-mm-dd, 숫자는 소수점 '.', 빈 값은 생략
'   UTF-8 BOM 1개 (Excel 한글 + pandas utf-8-sig)
'
' 실행:
'   1) Matching_data.xlsm 에서 ExportMatchingDataToCsv
'   2) 또는 ExportMatchingDataFromDefaultPath (아래 경로를 읽기 전용으로 연 뒤 저장하고 닫기)
'===============================================================================

Option Explicit

Private Const DEFAULT_WORKBOOK As String = "C:\Users\infomax\Documents\market_db_dashboard\Matching_data.xlsm"
Private Const OUTPUT_CSV_PATH As String = "C:\Users\infomax\Documents\market_db_dashboard\matching_data.csv"

Private Const SHEET_NAME As String = "Sheet1"
Private Const HEADER_SCAN_MAX_ROW As Long = 20
Private Const FALLBACK_HEADER_ROW As Long = 3
Private Const EXPORT_FLUSH_BATCH As Long = 20000

'-------------------------------------------------------------------------------

Public Sub ExportMatchingDataToCsv()
    On Error GoTo EH
    RunExport ActiveWorkbook, False, False
    Exit Sub
EH:
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

' 메시지 상자 없이 실행. 오류는 호출 쪽으로 다시 던진다. (스케줄/VBS 자동화용)
Public Sub ExportMatchingDataToCsvSilent()
    RunExport ActiveWorkbook, False, True
End Sub

Public Sub ExportMatchingDataFromDefaultPath()
    Dim wb As Workbook
    Dim openedHere As Boolean

    On Error GoTo EH

    Set wb = FindOpenWorkbook(DEFAULT_WORKBOOK)
    If wb Is Nothing Then
        If Dir$(DEFAULT_WORKBOOK) = "" Then
            MsgBox "파일 없음: " & DEFAULT_WORKBOOK, vbExclamation
            Exit Sub
        End If
        Set wb = Workbooks.Open(DEFAULT_WORKBOOK, ReadOnly:=True)
        openedHere = True
    End If

    RunExport wb, openedHere, False
    Exit Sub

EH:
    MsgBox "오류: " & Err.Description, vbCritical
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
    End If
End Sub

Private Sub RunExport(ByVal wb As Workbook, ByVal CloseWhenDone As Boolean, ByVal Silent As Boolean)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim dataStart As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim dateCols() As Long
    Dim nBlocks As Long
    Dim hdr As Variant
    Dim dat As Variant
    Dim outPath As String
    Dim stream As Object
    Dim batchLines() As String
    Dim batchCount As Long
    Dim numRows As Long
    Dim i As Long
    Dim savedScreen As Boolean
    Dim savedEvents As Boolean
    Dim savedCalc As XlCalculation
    Dim uiSaved As Boolean
    Dim failNum As Long
    Dim failDesc As String

    On Error GoTo EH

    savedScreen = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    savedCalc = Application.Calculation
    uiSaved = True

    Set ws = ResolveSheet(wb)
    If ws Is Nothing Then
        NoteFailure failNum, failDesc, Silent, "시트를 찾지 못했습니다."
        GoTo CleanExit
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True

    headerRow = FindHeaderRow(ws)
    If headerRow < 2 Then
        NoteFailure failNum, failDesc, Silent, "헤더 행에서 '일자' 열을 찾지 못했습니다."
        GoTo CleanExit
    End If
    dataStart = headerRow + 1

    lastCol = LastUsedCol(ws, headerRow)
    If lastCol < 2 Then
        NoteFailure failNum, failDesc, Silent, "헤더 열을 찾지 못했습니다."
        GoTo CleanExit
    End If

    dateCols = CollectDateCols(ws, headerRow, lastCol, nBlocks)
    If nBlocks < 1 Then
        NoteFailure failNum, failDesc, Silent, "헤더 행에서 '일자' 열을 찾지 못했습니다."
        GoTo CleanExit
    End If

    lastRow = MaxLastRowOnCols(ws, dateCols)
    If lastRow < dataStart Then
        NoteFailure failNum, failDesc, Silent, "데이터 행이 없습니다."
        GoTo CleanExit
    End If

    SetStatus "CSV: 시트 읽는 중..."
    hdr = ws.Range(ws.Cells(1, 1), ws.Cells(headerRow, lastCol)).Value2
    dat = ws.Range(ws.Cells(dataStart, 1), ws.Cells(lastRow, lastCol)).Value2

    outPath = OUTPUT_CSV_PATH
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText "date,instrument,field,value" & vbCrLf, 0

    numRows = 0
    batchCount = 0
    ReDim batchLines(0 To EXPORT_FLUSH_BATCH - 1)

    For i = 0 To nBlocks - 1
        AppendBlock stream, hdr, dat, headerRow, dateCols, i, nBlocks, lastCol, _
                    batchLines, batchCount, numRows
    Next i

    If batchCount > 0 Then FlushBatch stream, batchLines, batchCount

    SetStatus "CSV: 저장 중..."
    On Error Resume Next
    Kill outPath
    On Error GoTo EH
    stream.SaveToFile outPath, 2
    stream.Close
    Set stream = Nothing

    If Not Silent Then MsgBox "저장 완료: " & outPath & vbCrLf & "행 수: " & Format$(numRows, "#,##0"), vbInformation
    GoTo CleanExit

EH:
    failNum = Err.Number
    failDesc = Err.Description
    If Not Silent Then MsgBox "오류: " & failNum & " - " & failDesc, vbCritical

CleanExit:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    If uiSaved Then
        Application.StatusBar = False
        Application.Calculation = savedCalc
        Application.EnableEvents = savedEvents
        Application.ScreenUpdating = savedScreen
    End If
    On Error GoTo 0

CloseWb:
    If CloseWhenDone Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    If failNum <> 0 And Silent Then Err.Raise failNum, "MatchingData_ExportToCsv", failDesc
End Sub

Private Sub AppendBlock(ByVal stream As Object, ByVal hdr As Variant, ByVal dat As Variant, _
                        ByVal headerRow As Long, ByRef dateCols() As Long, ByVal blockIndex As Long, _
                        ByVal nBlocks As Long, ByVal lastCol As Long, _
                        ByRef batchLines() As String, ByRef batchCount As Long, ByRef numRows As Long)
    Dim dateCol As Long
    Dim blockEnd As Long
    Dim prevDateCol As Long
    Dim instrument As String
    Dim escapedInst As String
    Dim c As Long
    Dim r As Long
    Dim fieldName As String
    Dim escapedField As String
    Dim dateVal As String
    Dim cellVal As Variant
    Dim numText As String
    Dim line As String
    Dim nDataRows As Long

    dateCol = dateCols(blockIndex)
    If blockIndex < nBlocks - 1 Then
        blockEnd = dateCols(blockIndex + 1) - 1
    Else
        blockEnd = lastCol
    End If
    If blockEnd < dateCol Then blockEnd = dateCol

    prevDateCol = 0
    If blockIndex > 0 Then prevDateCol = dateCols(blockIndex - 1)

    instrument = BlockInstrument(hdr, headerRow, dateCol, blockEnd, prevDateCol, blockIndex + 1)
    escapedInst = EscapeCsv(instrument)

    nDataRows = UBound(dat, 1)
    For r = 1 To nDataRows
        dateVal = FormatDateForCsv(dat(r, dateCol))
        If Len(dateVal) = 0 Then GoTo NextDataRow

        For c = dateCol + 1 To blockEnd
            fieldName = FieldLabel(hdr, headerRow, c)
            If Len(fieldName) = 0 Then GoTo NextField
            cellVal = dat(r, c)
            If IsEmpty(cellVal) Or IsNull(cellVal) Or IsError(cellVal) Then GoTo NextField
            numText = FormatValueForCsv(cellVal)
            If Len(numText) = 0 Then GoTo NextField

            escapedField = EscapeCsv(fieldName)
            line = dateVal & "," & escapedInst & "," & escapedField & "," & numText
            batchLines(batchCount) = line
            batchCount = batchCount + 1
            numRows = numRows + 1
            If batchCount >= EXPORT_FLUSH_BATCH Then
                FlushBatch stream, batchLines, batchCount
                batchCount = 0
                SetStatus "CSV: " & Format$(numRows, "#,##0") & "행..."
                DoEvents
            End If
NextField:
        Next c
NextDataRow:
    Next r
End Sub

Private Function ResolveSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set ResolveSheet = wb.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ResolveSheet Is Nothing Then
        If wb.Worksheets.Count > 0 Then Set ResolveSheet = wb.Worksheets(1)
    End If
End Function

Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    Dim target As String
    target = LCase$(fullPath)
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = target Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet) As Long
    Dim lastCol As Long
    Dim rr As Long
    Dim cc As Long
    Dim bestRow As Long
    Dim bestCount As Long
    Dim cnt As Long
    Dim v As Variant
    Dim scan As Variant

    lastCol = ws.Cells(FALLBACK_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    scan = ws.Range(ws.Cells(1, 1), ws.Cells(HEADER_SCAN_MAX_ROW, lastCol)).Value2
    bestRow = 0
    bestCount = 0

    For rr = 1 To HEADER_SCAN_MAX_ROW
        cnt = 0
        For cc = 1 To lastCol
            v = scan(rr, cc)
            If Not IsError(v) Then
                If Trim$(CStr(v)) = "일자" Then cnt = cnt + 1
            End If
        Next cc
        If cnt > bestCount Then
            bestCount = cnt
            bestRow = rr
        End If
    Next rr

    If bestCount < 1 Then
        FindHeaderRow = 0
    Else
        FindHeaderRow = bestRow
    End If
End Function

Private Function LastUsedCol(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim best As Long
    Dim tmp As Long
    Dim rr As Long

    best = 1
    For rr = headerRow - 1 To headerRow
        If rr >= 1 Then
            tmp = ws.Cells(rr, ws.Columns.Count).End(xlToLeft).Column
            If tmp > best Then best = tmp
        End If
    Next rr
    LastUsedCol = best
End Function

Private Function CollectDateCols(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastCol As Long, ByRef nBlocks As Long) As Long()
    Dim cols() As Long
    Dim c As Long
    Dim v As Variant

    nBlocks = 0
    ReDim cols(0 To 0)
    For c = 1 To lastCol
        v = ws.Cells(headerRow, c).Value2
        If Not IsError(v) Then
            If Trim$(CStr(v)) = "일자" Then
                If nBlocks > 0 Then ReDim Preserve cols(0 To nBlocks)
                cols(nBlocks) = c
                nBlocks = nBlocks + 1
            End If
        End If
    Next c
    CollectDateCols = cols
End Function

Private Function MaxLastRowOnCols(ByVal ws As Worksheet, ByRef dateCols() As Long) As Long
    Dim i As Long
    Dim tmp As Long
    Dim best As Long

    best = 0
    For i = LBound(dateCols) To UBound(dateCols)
        tmp = ws.Cells(ws.Rows.Count, dateCols(i)).End(xlUp).Row
        If tmp > best Then best = tmp
    Next i
    MaxLastRowOnCols = best
End Function

Private Function BlockInstrument(ByVal hdr As Variant, ByVal headerRow As Long, ByVal dateCol As Long, _
                                 ByVal blockEnd As Long, ByVal prevDateCol As Long, ByVal blockNo As Long) As String
    Dim rr As Long
    Dim c As Long
    Dim s As String

    For rr = headerRow - 1 To 1 Step -1
        For c = dateCol To prevDateCol + 1 Step -1
            s = HeaderText(hdr, rr, c)
            If Len(s) > 0 Then
                If s <> "일자" And Not IsMetaLabel(s) And Not LooksLikeIsoDate(s) Then
                    BlockInstrument = s
                    Exit Function
                End If
            End If
        Next c
    Next rr

    For rr = headerRow - 1 To 1 Step -1
        For c = dateCol To blockEnd
            s = HeaderText(hdr, rr, c)
            If Len(s) > 0 Then
                If s <> "일자" And Not IsMetaLabel(s) And Not LooksLikeIsoDate(s) Then
                    BlockInstrument = s
                    Exit Function
                End If
            End If
        Next c
    Next rr

    BlockInstrument = "Block" & CStr(blockNo)
End Function

Private Function FieldLabel(ByVal hdr As Variant, ByVal headerRow As Long, ByVal col As Long) As String
    Dim s As String
    s = HeaderText(hdr, headerRow, col)
    If Len(s) = 0 Then Exit Function
    If s = "일자" Then Exit Function
    If LooksLikeIsoDate(s) Then Exit Function
    FieldLabel = s
End Function

Private Function HeaderText(ByVal hdr As Variant, ByVal row As Long, ByVal col As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = hdr(row, col)
    On Error GoTo 0
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    HeaderText = Trim$(CStr(v))
End Function

Private Function IsMetaLabel(ByVal s As String) As Boolean
    Select Case s
        Case "시작", "종료", "Data 개수", "주기", "정렬", "영업일", "시세산출"
            IsMetaLabel = True
        Case Else
            IsMetaLabel = False
    End Select
End Function

Private Function FormatDateForCsv(ByVal v As Variant) As String
    Dim serial As Double
    Dim s As String

    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function

    If IsNumeric(v) Then
        serial = CDbl(v)
        If serial >= 1# And serial < 2958466# Then
            FormatDateForCsv = Format$(CDate(serial), "yyyy-mm-dd")
            Exit Function
        End If
    End If

    s = Trim$(CStr(v))
    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        FormatDateForCsv = s
        Exit Function
    End If

    On Error Resume Next
    FormatDateForCsv = Format$(CDate(v), "yyyy-mm-dd")
    If Err.Number <> 0 Then
        Err.Clear
        FormatDateForCsv = vbNullString
    End If
    On Error GoTo 0
End Function

Private Function FormatValueForCsv(ByVal v As Variant) As String
    Dim d As Double
    Dim s As String
    Dim decSep As String
    Dim dotPos As Long

    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function

    ' 금리(3.818 등)는 엑셀 날짜 일련번호 범위에 걸치므로 IsDate 로 거르면 안 된다.
    If VarType(v) <> vbString And IsNumeric(v) Then
        d = CDbl(v)
        s = Format$(d, "0.00000000")
        decSep = Application.International(xlDecimalSeparator)
        If decSep <> "." Then s = Replace$(s, decSep, ".")
        dotPos = InStr(s, ".")
        If dotPos > 0 Then
            Do While Len(s) > dotPos And Right$(s, 1) = "0"
                s = Left$(s, Len(s) - 1)
            Loop
            If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
        End If
        FormatValueForCsv = s
        Exit Function
    End If

    FormatValueForCsv = EscapeCsv(Trim$(CStr(v)))
End Function

Private Function LooksLikeIsoDate(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) <> 10 Then Exit Function
    If Mid$(t, 5, 1) <> "-" Or Mid$(t, 8, 1) <> "-" Then Exit Function
    LooksLikeIsoDate = IsNumeric(Left$(t, 4)) And IsNumeric(Mid$(t, 6, 2)) And IsNumeric(Right$(t, 2))
End Function

Private Function EscapeCsv(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If InStr(1, t, ",") > 0 Or InStr(1, t, """") > 0 Or InStr(1, t, vbLf) > 0 Or InStr(1, t, vbCr) > 0 Then
        EscapeCsv = """" & Replace$(t, """", """""") & """"
    Else
        EscapeCsv = t
    End If
End Function

Private Sub FlushBatch(ByVal stream As Object, ByRef lines() As String, ByVal n As Long)
    Dim i As Long
    Dim chunk() As String
    If n <= 0 Then Exit Sub
    ReDim chunk(0 To n - 1)
    For i = 0 To n - 1
        chunk(i) = lines(i)
    Next i
    stream.WriteText Join(chunk, vbCrLf) & vbCrLf, 0
End Sub

Private Sub SetStatus(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Private Sub NoteFailure(ByRef failNum As Long, ByRef failDesc As String, ByVal Silent As Boolean, ByVal msg As String)
    failNum = vbObjectError + 513
    failDesc = msg
    If Not Silent Then MsgBox msg, vbExclamation
End Sub
