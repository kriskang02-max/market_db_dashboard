' ============================================================
' fund_db.xlsm → long CSV (fund, field, date, value)
'
' 시트 구조 (가로 반복 블록):
'   행1: 메타(시작/종료 등)
'   행2: 각 블록의 첫 열(A, F, K, …)에 펀드명 (병합 가능)
'   행3: 각 블록마다 "일자" | 기준가 | 수정기준가 | … (블록 너비는 가변, 다음 "일자" 열까지)
'   행4~: 해당 블록의 일자 열 + 값 열들
'
' 블록 탐지: 헤더 행에서 셀 값이 "일자"인 열 = 블록 시작 (열 추가·펀드 개수 증가 자동 반영)
' 성능: Range.Value2 일괄 읽기, Shell 정렬, Join
' ============================================================

Option Explicit

Private Const FUND_SHEET_NAME As String = ""
Private Const FUND_SHEET_INDEX As Long = 1

Private Const DATE_COL As Long = 1

Private Const AUTO_DETECT_ILJA_HEADER_ROW As Boolean = True
Private Const ILJA_HEADER_SCAN_MAX_ROW As Long = 40
Private Const FALLBACK_HEADER_ROW As Long = 3

Private Const FUND_DB_CSV_PATH As String = "C:\Users\infomax\Documents\market_db_dashboard\fund_db.csv"

Private Const SKIP_WAIT_FOR_LOAD As Boolean = True
Private Const IMDH_READY_CELL_ADDRESSES As String = "A2,F2,K2,P2"
Private Const IMDH_WAIT_TIMEOUT_SECONDS As Long = 300
Private Const IMDH_POLL_SECONDS As Long = 5

' ----------------------------------------------------------------
Public Sub ExportFundDbToCsv()
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim fundRow As Long
    Dim dataStartRow As Long
    Dim lastCol As Long, lastRow As Long
    Dim nDataRows As Long
    Dim hdr As Variant, dat As Variant
    Dim fieldNames() As String
    Dim blockStarts() As Long
    Dim nBlocks As Long
    Dim bi As Long
    Dim cStart As Long, cEnd As Long
    Dim fundName As String
    Dim ri As Long, ci As Long
    Dim dateOut As String
    Dim numRows As Long
    Dim maxOut As Long
    Dim funds() As String, fields() As String, dtes() As String, vals() As String
    Dim lines() As String
    Dim stream As Object
    Dim k As Long

    On Error GoTo ErrHandle

    Set ws = ResolveFundSheet()
    If ws Is Nothing Then
        MsgBox "시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    If Not WaitForImdhReady(ws, IMDH_READY_CELL_ADDRESSES, IMDH_WAIT_TIMEOUT_SECONDS, IMDH_POLL_SECONDS) Then
        MsgBox "IMDH 데이터가 준비되지 않았습니다. (" & IMDH_READY_CELL_ADDRESSES & ")", vbExclamation
        Exit Sub
    End If

    ResolveFundLayout ws, headerRow, fundRow, dataStartRow

    lastCol = LastNonEmptyColInRow(ws, headerRow)
    If lastCol < 1 Then
        MsgBox "헤더 행에서 열을 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    blockStarts = CollectIljaBlockStarts(ws, headerRow, lastCol, nBlocks)
    If nBlocks < 1 Then
        MsgBox """일자"" 열(블록 시작)을 헤더 행에서 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = MaxLastRowAcrossBlockStarts(ws, blockStarts, nBlocks)
    If lastRow < dataStartRow Then
        MsgBox "데이터 행이 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not SKIP_WAIT_FOR_LOAD Then
        OptionalRecalcFundSheet ws
    End If

    hdr = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol)).Value2
    fieldNames = FieldNamesFromHeaderRow(hdr, lastCol)

    nDataRows = lastRow - dataStartRow + 1
    dat = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(lastRow, lastCol)).Value2

    If Not Is2DArray(dat) Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "데이터 범위를 배열로 읽지 못했습니다.", vbExclamation
        Exit Sub
    End If

    maxOut = nDataRows * (lastCol + 4)
    If maxOut < 32 Then maxOut = 32
    ReDim funds(1 To maxOut)
    ReDim fields(1 To maxOut)
    ReDim dtes(1 To maxOut)
    ReDim vals(1 To maxOut)

    numRows = 0
    For bi = 1 To nBlocks
        cStart = blockStarts(bi)
        If bi < nBlocks Then
            cEnd = blockStarts(bi + 1) - 1
        Else
            cEnd = lastCol
        End If

        fundName = FundNameFromRowAtCol(ws, fundRow, cStart)
        If Len(fundName) = 0 Then
            fundName = "Fund_Block_" & CStr(cStart)
        End If

        For ri = 1 To UBound(dat, 1)
            If Not CellIsLikelyDateValue(dat(ri, cStart)) Then GoTo NextRi
            dateOut = NormalizeDateForCsv(dat(ri, cStart))
            If Len(dateOut) = 0 Then GoTo NextRi

            For ci = cStart + 1 To cEnd
                If Len(fieldNames(ci)) = 0 Then GoTo NextCi
                If StrComp(fieldNames(ci), "일자", vbTextCompare) = 0 Then GoTo NextCi

                If IsEmpty(dat(ri, ci)) Or IsError(dat(ri, ci)) Then GoTo NextCi

                numRows = numRows + 1
                If numRows > maxOut Then
                    maxOut = maxOut * 2
                    ReDim Preserve funds(1 To maxOut)
                    ReDim Preserve fields(1 To maxOut)
                    ReDim Preserve dtes(1 To maxOut)
                    ReDim Preserve vals(1 To maxOut)
                End If
                funds(numRows) = fundName
                fields(numRows) = fieldNames(ci)
                dtes(numRows) = dateOut
                vals(numRows) = NormalizeValueForCsv(dat(ri, ci))
NextCi:
            Next ci
NextRi:
        Next ri
    Next bi

    If numRows = 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "보낼 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    ShellSort4 funds, fields, dtes, vals, numRows

    ReDim lines(1 To numRows)
    For k = 1 To numRows
        lines(k) = EscapeCsv(funds(k)) & "," & EscapeCsv(fields(k)) & "," & EscapeCsv(dtes(k)) & "," & EscapeCsv(vals(k))
    Next k

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText ChrW(65279), 0
    stream.WriteText "fund,field,date,value" & vbCrLf, 0
    stream.WriteText Join(lines, vbCrLf), 0

    On Error Resume Next
    Kill FUND_DB_CSV_PATH
    On Error GoTo ErrHandle

    stream.SaveToFile FUND_DB_CSV_PATH, 2
    stream.Close
    Set stream = Nothing

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "저장 완료 (펀드 블록 " & nBlocks & "개, 헤더행=" & CStr(headerRow) & ")" & vbCrLf & FUND_DB_CSV_PATH & vbCrLf & "행 수: " & numRows, vbInformation
    Exit Sub

ErrHandle:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

Private Function WaitForImdhReady(ByVal ws As Worksheet, ByVal csvAddresses As String, ByVal timeoutSeconds As Long, ByVal pollSeconds As Long) As Boolean
    Dim startedAt As Single
    Dim elapsed As Single
    Dim parts() As String
    Dim i As Long
    Dim addr As String
    Dim v As Variant
    Dim allReady As Boolean

    If pollSeconds < 1 Then pollSeconds = 1
    startedAt = Timer
    parts = Split(csvAddresses, ",")

    Do
        On Error Resume Next
        Application.Calculate
        On Error GoTo 0

        allReady = True
        For i = LBound(parts) To UBound(parts)
            addr = Trim$(parts(i))
            If Len(addr) = 0 Then GoTo NextAddr

            On Error Resume Next
            v = ws.Range(addr).Value2
            If Err.Number <> 0 Then
                Err.Clear
                allReady = False
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0

            If IsError(v) Then
                allReady = False
                Exit For
            End If
            If Len(Trim$(CStr(v))) = 0 Then
                allReady = False
                Exit For
            End If
NextAddr:
        Next i

        If allReady Then
            WaitForImdhReady = True
            Exit Function
        End If

        elapsed = Timer - startedAt
        If elapsed < 0 Then elapsed = elapsed + 86400
        If elapsed >= timeoutSeconds Then Exit Do

        Application.Wait Now + TimeSerial(0, 0, pollSeconds)
        DoEvents
    Loop

    WaitForImdhReady = False
End Function

' 헤더 행에서 "일자"인 열 = 각 펀드 블록의 시작 열
Private Function CollectIljaBlockStarts(ByVal ws As Worksheet, ByVal hr As Long, ByVal lastCol As Long, ByRef nOut As Long) As Long()
    Dim tmp() As Long
    Dim c As Long, n As Long
    ReDim tmp(1 To 64)
    n = 0
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(hr, c).Value2)), "일자", vbTextCompare) = 0 Then
            n = n + 1
            If n > UBound(tmp) Then ReDim Preserve tmp(1 To UBound(tmp) * 2)
            tmp(n) = c
        End If
    Next c
    nOut = n
    If n = 0 Then
        CollectIljaBlockStarts = tmp
        Exit Function
    End If
    ReDim Preserve tmp(1 To n)
    CollectIljaBlockStarts = tmp
End Function

Private Function MaxLastRowAcrossBlockStarts(ByVal ws As Worksheet, ByRef blockStarts() As Long, ByVal nBlocks As Long) As Long
    Dim i As Long
    Dim lr As Long
    Dim m As Long
    m = 0
    For i = 1 To nBlocks
        lr = ws.Cells(ws.Rows.Count, blockStarts(i)).End(xlUp).Row
        If lr > m Then m = lr
    Next i
    MaxLastRowAcrossBlockStarts = m
End Function

Private Function Is2DArray(ByVal v As Variant) As Boolean
    On Error Resume Next
    Is2DArray = (UBound(v, 1) >= LBound(v, 1)) And (UBound(v, 2) >= LBound(v, 2))
    On Error GoTo 0
End Function

Private Function FieldNamesFromHeaderRow(ByVal hdr As Variant, ByVal lastCol As Long) As String()
    Dim out() As String
    Dim c As Long
    ReDim out(1 To lastCol)
    If Not IsArray(hdr) Then
        out(1) = Trim$(CStr(hdr))
        FieldNamesFromHeaderRow = out
        Exit Function
    End If
    On Error Resume Next
    If UBound(hdr, 1) = 1 Then
        For c = 1 To lastCol
            out(c) = Trim$(CStr(hdr(1, c)))
        Next c
    ElseIf UBound(hdr, 2) = 1 Then
        For c = 1 To lastCol
            out(c) = Trim$(CStr(hdr(c, 1)))
        Next c
    Else
        For c = 1 To lastCol
            out(c) = Trim$(CStr(hdr(1, c)))
        Next c
    End If
    On Error GoTo 0
    FieldNamesFromHeaderRow = out
End Function

Private Sub ShellSort4(ByRef f() As String, ByRef fld() As String, ByRef dte() As String, ByRef valStr() As String, ByVal n As Long)
    Dim gap As Long, i As Long, j As Long
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            j = i
            Do While j > gap
                If CompareFundFieldDate(f(j - gap), fld(j - gap), dte(j - gap), f(j), fld(j), dte(j)) <= 0 Then Exit Do
                Swap4 f, fld, dte, valStr, j - gap, j
                j = j - gap
            Loop
        Next i
        gap = gap \ 2
    Loop
End Sub

Private Sub Swap4(ByRef f() As String, ByRef fld() As String, ByRef dte() As String, ByRef valStr() As String, ByVal i As Long, ByVal j As Long)
    Dim tf As String, tFld As String, td As String, tv As String
    tf = f(i): f(i) = f(j): f(j) = tf
    tFld = fld(i): fld(i) = fld(j): fld(j) = tFld
    td = dte(i): dte(i) = dte(j): dte(j) = td
    tv = valStr(i): valStr(i) = valStr(j): valStr(j) = tv
End Sub

Private Sub ResolveFundLayout(ByVal ws As Worksheet, ByRef headerRow As Long, ByRef fundRow As Long, ByRef dataStartRow As Long)
    headerRow = 0
    If AUTO_DETECT_ILJA_HEADER_ROW Then headerRow = FindIljaHeaderRow(ws)
    If headerRow < 1 Then headerRow = FALLBACK_HEADER_ROW
    fundRow = headerRow - 1
    If fundRow < 1 Then fundRow = 1
    dataStartRow = headerRow + 1
End Sub

' A열에 "일자"가 있는 행 = 필드 헤더 행 (첫 블록 기준)
Private Function FindIljaHeaderRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim t As String
    FindIljaHeaderRow = 0
    For r = 1 To ILJA_HEADER_SCAN_MAX_ROW
        t = Trim$(CStr(ws.Cells(r, DATE_COL).Value2))
        If StrComp(t, "일자", vbTextCompare) = 0 Then
            FindIljaHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function LastNonEmptyColInRow(ByVal ws As Worksheet, ByVal rr As Long) As Long
    LastNonEmptyColInRow = ws.Cells(rr, ws.Columns.Count).End(xlToLeft).Column
    If LastNonEmptyColInRow < 1 Then LastNonEmptyColInRow = 1
End Function

Private Function CellIsLikelyDateValue(ByVal v As Variant) As Boolean
    Dim t As String
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        t = Trim$(CStr(v))
        If Len(t) = 0 Then Exit Function
        If StrComp(t, "일자", vbTextCompare) = 0 Then Exit Function
        If StrComp(t, "시작", vbTextCompare) = 0 Then Exit Function
        If StrComp(t, "종료", vbTextCompare) = 0 Then Exit Function
    End If
    If IsDate(v) Then CellIsLikelyDateValue = True: Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000# And CDbl(v) < 60000# Then
            On Error Resume Next
            CellIsLikelyDateValue = IsDate(CDate(CDbl(v)))
            On Error GoTo 0
        End If
    End If
End Function

Private Function CompareFundFieldDate(ByVal f1 As String, ByVal fld1 As String, ByVal d1 As String, ByVal f2 As String, ByVal fld2 As String, ByVal d2 As String) As Long
    Dim x As Long
    x = StrComp(f1, f2, vbTextCompare)
    If x <> 0 Then CompareFundFieldDate = x: Exit Function
    x = StrComp(fld1, fld2, vbTextCompare)
    If x <> 0 Then CompareFundFieldDate = x: Exit Function
    CompareFundFieldDate = StrComp(d1, d2, vbTextCompare)
End Function

Private Function ResolveFundSheet() As Worksheet
    On Error Resume Next
    If Len(Trim$(FUND_SHEET_NAME)) > 0 Then
        Set ResolveFundSheet = ThisWorkbook.Worksheets(Trim$(FUND_SHEET_NAME))
    Else
        If ThisWorkbook.Worksheets.Count >= FUND_SHEET_INDEX And FUND_SHEET_INDEX >= 1 Then
            Set ResolveFundSheet = ThisWorkbook.Worksheets(FUND_SHEET_INDEX)
        End If
    End If
    On Error GoTo 0
End Function

' 블록 첫 열(fundRow, colIdx)에서 펀드명 — 병합이면 영역의 왼쪽 위
Private Function FundNameFromRowAtCol(ByVal ws As Worksheet, ByVal fundRow As Long, ByVal colIdx As Long) As String
    Dim rng As Range
    Dim t As String
    On Error Resume Next
    Set rng = ws.Cells(fundRow, colIdx)
    If rng.MergeCells Then Set rng = rng.MergeArea.Cells(1, 1)
    t = Trim$(CStr(rng.Value2))
    If Len(t) = 0 Then t = Trim$(CStr(rng.Text))
    On Error GoTo 0
    FundNameFromRowAtCol = t
End Function

Private Function NormalizeDateForCsv(ByVal v As Variant) As String
    If IsEmpty(v) Or IsError(v) Then NormalizeDateForCsv = vbNullString: Exit Function
    If IsDate(v) Then
        NormalizeDateForCsv = Format$(CDate(v), "yyyy-mm-dd")
        Exit Function
    End If
    If IsNumeric(v) Then
        If CDbl(v) > 20000# And CDbl(v) < 60000# Then
            On Error Resume Next
            NormalizeDateForCsv = Format$(CDate(CDbl(v)), "yyyy-mm-dd")
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            If Len(NormalizeDateForCsv) > 0 Then Exit Function
        End If
    End If
    If VarType(v) = vbString Then
        If IsDate(v) Then
            NormalizeDateForCsv = Format$(CDate(v), "yyyy-mm-dd")
            Exit Function
        End If
    End If
    NormalizeDateForCsv = Trim$(CStr(v))
End Function

Private Function NormalizeValueForCsv(ByVal v As Variant) As String
    If IsError(v) Then NormalizeValueForCsv = CStr(v): Exit Function
    If IsNumeric(v) Then
        NormalizeValueForCsv = Replace(Format$(CDbl(v), "0.##############"), ",", ".")
    Else
        NormalizeValueForCsv = Trim$(CStr(v))
    End If
End Function

Private Function EscapeCsv(ByVal s As String) As String
    Dim t As String
    t = s
    If InStr(1, t, ",") > 0 Or InStr(1, t, """") > 0 Or InStr(1, t, vbLf) > 0 Or InStr(1, t, vbCr) > 0 Then
        EscapeCsv = """" & Replace(t, """", """""") & """"
    Else
        EscapeCsv = t
    End If
End Function

Private Sub OptionalRecalcFundSheet(ByVal ws As Worksheet)
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    ws.Calculate
    On Error GoTo 0
End Sub
