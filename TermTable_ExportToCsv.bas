Attribute VB_Name = "TermTable_ExportToCsv"
'===============================================================================
' term_table.xlsx -> long CSV (CP term + Bond term)
'
' CP term structure
'   Col K (11)     : CP 발행사명 — 금리 열(L~AO)과 반드시 같은 행 (L49 금리 ↔ K49 발행사)
'   Col E (5)·F(6) : E에 있는 발행사명으로 F의 등급을 조회(동일 행이 아닐 수 있음 → K 발행사명으로 E열 검색)
'   Col L:U (12-21): latest-day rates (one as-of date per col in row 2, tenor row 3)
'   Col V:AE (22-31): T-1 rates
'   Col AF:AO (32-41): T-2 rates
'   Row 2          : as-of date for that column (merged OK -> walk left)
'   Row 3          : maturity label for that column
'   Row 4+         : rate values
'
' Bond term structure
'   Col AP (42)    : issuer name (row 4+) — CSV에는 접두어 "시가평가 4사평균 " 제거
'   Col AS (45)    : 문자열에서 마지막 "채" 또는 마지막 "증" 중 더 뒤에 오는 글자 다음부터 등급 토큰 추출
'   Col AU:BI (47-61): latest rates; row 3 = maturity, row 2 = as-of if present
'   Row 4+         : values
'
' Output: term_table_long.csv (UTF-8 BOM)
'   section,날짜,발행사명,등급,만기,금리,섹터
'   CP 섹터: 지주·증권·카드·여전·일반 (발행사명 키워드, 현대커머셜→여전)
'   채권 섹터: 특수·은행·카드·여전·일반 (AS열 전체 문자열 키워드)
'===============================================================================

Option Explicit

Private Const SHEET_NAME As String = "Sheet1"

Private Const HEADER_ASOF_ROW As Long = 2   ' 기준일
Private Const HEADER_TENOR_ROW As Long = 3  ' 만기
Private Const DATA_FIRST_ROW As Long = 4    ' 첫 데이터 행

' 메타: 마지막 데이터 행 잡을 때만 사용 (A열)
Private Const ROW_META_FIRST_COL As Long = 1   ' A

' CP: 금리(L~)와 같은 행의 발행사명 열 (보통 K=11; 시트가 J열이면 10으로 변경)
Private Const COL_CP_ISSUER As Long = 11       ' K
' 등급: E열 발행사명 = 키, F열 = 값 (K열 발행사명으로 E를 스캔해 F 반환)
Private Const COL_CP_LOOKUP_ISSUER As Long = 5  ' E
Private Const COL_CP_LOOKUP_RATING As Long = 6 ' F
Private Const COL_CP_FIRST As Long = 12        ' L
Private Const COL_CP_LAST As Long = 41         ' AO
Private Const COL_CP_LATEST_LAST As Long = 21  ' U
Private Const COL_CP_T1_LAST As Long = 31      ' AE

' Bond
Private Const COL_BD_ISSUER As Long = 42       ' AP
Private Const COL_BD_RATING_SRC As Long = 45   ' AS (무보증 뒤 등급 파싱)
Private Const COL_BD_FIRST As Long = 47        ' AU
Private Const COL_BD_LAST As Long = 61         ' BI (끝 열 바뀌면 여만 수정)

Private Const OUT_FILE_NAME As String = "term_table_long.csv"

'-------------------------------------------------------------------------------

Public Sub ExportTermTableToLongCsv()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outPath As String
    Dim lastDataRow As Long
    Dim r As Long, c As Long
    Dim section As String
    Dim dt As String, tenor As String
    Dim issuer As String
    Dim rating As String
    Dim sector As String
    Dim v As Variant
    Dim lines() As String
    Dim n As Long
    Dim cap As Long
    Dim cpRatingMap As Object

    cap = 16384
    ReDim lines(1 To cap)
    n = 0

    On Error GoTo EH

    Set wb = ActiveWorkbook
    If Len(wb.Path) = 0 Then
        MsgBox "Save the workbook to a folder first, then run again.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_NAME)
    On Error GoTo EH
    If ws Is Nothing Then
        MsgBox "Worksheet not found: " & SHEET_NAME, vbExclamation
        Exit Sub
    End If

    lastDataRow = GetLastDataRow(ws)
    If lastDataRow < DATA_FIRST_ROW Then
        MsgBox "No data rows (need rows from " & DATA_FIRST_ROW & ").", vbExclamation
        Exit Sub
    End If

    Set cpRatingMap = BuildCpRatingLookupEf(ws, lastDataRow)

    n = n + 1
    lines(n) = JoinCsvRow(Array("section", "날짜", "발행사명", "등급", "만기", "금리", "섹터"))

    section = "cp_term"
    For c = COL_CP_FIRST To COL_CP_LAST
        dt = EffectiveAsOfYmd(ws, c)
        tenor = Trim$(Nz(ws.Cells(HEADER_TENOR_ROW, c).Value2))
        For r = DATA_FIRST_ROW To lastDataRow
            issuer = Trim$(Nz(ws.Cells(r, COL_CP_ISSUER).Value2))
            rating = CpRatingFromLookupMap(cpRatingMap, issuer)
            sector = CpSectorFromIssuer(issuer)
            v = ws.Cells(r, c).Value2
            If Not IsEmpty(v) Then
                n = n + 1
                If n > UBound(lines) Then
                    cap = cap * 2
                    ReDim Preserve lines(1 To cap)
                End If
                lines(n) = JoinCsvRow(Array(section, dt, issuer, rating, tenor, FormatRate(v), sector))
            End If
        Next r
    Next c

    section = "bond_term"
    For c = COL_BD_FIRST To COL_BD_LAST
        dt = EffectiveAsOfYmd(ws, c)
        tenor = Trim$(Nz(ws.Cells(HEADER_TENOR_ROW, c).Value2))
        For r = DATA_FIRST_ROW To lastDataRow
            issuer = NormalizeBondIssuerName(Trim$(Nz(ws.Cells(r, COL_BD_ISSUER).Value2)))
            rating = BondRatingFromAsCell(ws.Cells(r, COL_BD_RATING_SRC).Value2)
            sector = BondSectorFromAsText(ws.Cells(r, COL_BD_RATING_SRC).Value2)
            v = ws.Cells(r, c).Value2
            If Not IsEmpty(v) Then
                n = n + 1
                If n > UBound(lines) Then
                    cap = cap * 2
                    ReDim Preserve lines(1 To cap)
                End If
                lines(n) = JoinCsvRow(Array(section, dt, issuer, rating, tenor, FormatRate(v), sector))
            End If
        Next r
    Next c

    outPath = wb.Path & Application.PathSeparator & OUT_FILE_NAME
    WriteUtf8Csv outPath, lines, n

    MsgBox "Saved: " & outPath & vbCrLf & "Rows (excl. header): " & (n - 1), vbInformation
    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub ExportTermTableFromDefaultPath()
    Const DEFAULT_XLSX As String = "C:\Users\infomax\Documents\market_db_dashboard\term_table.xlsx"
    Dim wb As Workbook

    If Dir(DEFAULT_XLSX) = "" Then
        MsgBox "File not found: " & DEFAULT_XLSX, vbExclamation
        Exit Sub
    End If

    Set wb = Nothing
    On Error GoTo FailOpen
    Set wb = Workbooks.Open(DEFAULT_XLSX, ReadOnly:=True)
    wb.Activate
    ExportTermTableToLongCsv
    GoTo CloseWb
FailOpen:
    MsgBox "Could not open: " & DEFAULT_XLSX & vbCrLf & Err.Description, vbCritical
CloseWb:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

' 2행 기준일(병합 시 왼쪽 앵커), yyyy-mm-dd 로 통일
Private Function EffectiveAsOfYmd(ByVal ws As Worksheet, ByVal c As Long) As String
    Dim i As Long
    Dim v As Variant
    For i = c To 1 Step -1
        v = ws.Cells(HEADER_ASOF_ROW, i).Value2
        If Not IsEmpty(v) Then
            If Len(Trim$(CStr(v))) > 0 Then
                EffectiveAsOfYmd = FormatDateYmd(v)
                Exit Function
            End If
        End If
    Next i
    EffectiveAsOfYmd = ""
End Function

Private Function FormatDateYmd(ByVal v As Variant) As String
    On Error Resume Next
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000# And CDbl(v) < 1000000# Then
            FormatDateYmd = Format$(CDate(v), "yyyy-mm-dd")
            Exit Function
        End If
    End If
    If IsDate(v) Then
        FormatDateYmd = Format$(CDate(v), "yyyy-mm-dd")
    Else
        FormatDateYmd = Trim$(CStr(v))
    End If
End Function

' E열 발행사 → F열 등급 맵 (같은 E 이름이 여러 행이면 아래쪽 행이 우선)
Private Function BuildCpRatingLookupEf(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim d As Object
    Dim r As Long
    Dim k As String
    Set d = CreateObject("Scripting.Dictionary")
    For r = DATA_FIRST_ROW To lastRow
        k = Trim$(Nz(ws.Cells(r, COL_CP_LOOKUP_ISSUER).Value2))
        If Len(k) > 0 Then
            d(k) = Trim$(Nz(ws.Cells(r, COL_CP_LOOKUP_RATING).Value2))
        End If
    Next r
    Set BuildCpRatingLookupEf = d
End Function

Private Function CpRatingFromLookupMap(ByVal d As Object, ByVal issuerK As String) As String
    If d Is Nothing Then Exit Function
    If Len(issuerK) = 0 Then Exit Function
    If d.Exists(issuerK) Then CpRatingFromLookupMap = CStr(d(issuerK))
End Function

' CP 섹터: 지주 / 증권 / 카드 / 여전 / 일반
Private Function CpSectorFromIssuer(ByVal issuerName As String) As String
    Dim s As String
    s = Trim$(issuerName)
    If Len(s) = 0 Then
        CpSectorFromIssuer = "일반"
        Exit Function
    End If
    If InStr(1, s, "현대커머셜", vbTextCompare) > 0 Then
        CpSectorFromIssuer = "여전"
        Exit Function
    End If
    If InStr(1, s, "지주", vbTextCompare) > 0 Then
        CpSectorFromIssuer = "지주"
        Exit Function
    End If
    If InStr(1, s, "증권", vbTextCompare) > 0 Then
        CpSectorFromIssuer = "증권"
        Exit Function
    End If
    If InStr(1, s, "카드", vbTextCompare) > 0 Then
        CpSectorFromIssuer = "카드"
        Exit Function
    End If
    If InStr(1, s, "캐피탈", vbTextCompare) > 0 Then
        CpSectorFromIssuer = "여전"
        Exit Function
    End If
    CpSectorFromIssuer = "일반"
End Function

' 채권 섹터: AS열 전체 문자열 기준 — 특수 / 은행 / 카드 / 여전 / 일반
Private Function BondSectorFromAsText(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(Nz(v))
    If Len(s) = 0 Then
        BondSectorFromAsText = "일반"
        Exit Function
    End If
    If InStr(1, s, "공사", vbTextCompare) > 0 Or InStr(1, s, "공단채", vbTextCompare) > 0 Then
        BondSectorFromAsText = "특수"
        Exit Function
    End If
    If InStr(1, s, "은행채", vbTextCompare) > 0 Then
        BondSectorFromAsText = "은행"
        Exit Function
    End If
    If InStr(1, s, "카드채", vbTextCompare) > 0 Then
        BondSectorFromAsText = "카드"
        Exit Function
    End If
    If InStr(1, s, "기타금융채", vbTextCompare) > 0 Then
        BondSectorFromAsText = "여전"
        Exit Function
    End If
    BondSectorFromAsText = "일반"
End Function

' AP 발행사명에서 "시가평가 4사평균 " 또는 "시가평가 4사평균" 접두 제거
Private Function NormalizeBondIssuerName(ByVal s As String) As String
    Dim t As String
    Const PFX_SPACE As String = "시가평가 4사평균 "
    Const PFX_TIGHT As String = "시가평가 4사평균"
    t = Trim$(s)
    If Len(t) = 0 Then
        NormalizeBondIssuerName = ""
        Exit Function
    End If
    If Len(t) >= Len(PFX_SPACE) Then
        If StrComp(Left$(t, Len(PFX_SPACE)), PFX_SPACE, vbTextCompare) = 0 Then
            NormalizeBondIssuerName = Trim$(Mid$(t, Len(PFX_SPACE) + 1))
            Exit Function
        End If
    End If
    If Len(t) >= Len(PFX_TIGHT) Then
        If StrComp(Left$(t, Len(PFX_TIGHT)), PFX_TIGHT, vbTextCompare) = 0 Then
            NormalizeBondIssuerName = Trim$(Mid$(t, Len(PFX_TIGHT) + 1))
            Exit Function
        End If
    End If
    NormalizeBondIssuerName = t
End Function

' AS 셀: "채" 또는 "증" 뒤에서 등급 시작 — 두 글자가 모두 있으면 문자열상 더 뒤 위치를 기준
Private Function BondRatingFromAsCell(ByVal v As Variant) As String
    Dim s As String
    Dim pChae As Long
    Dim pJeung As Long
    Dim p As Long
    Dim i As Long
    Dim ch As String
    Dim buf As String

    s = Trim$(Nz(v))
    If Len(s) = 0 Then Exit Function

    pChae = LastInStr1(s, "채")
    pJeung = LastInStr1(s, "증")
    p = pChae
    If pJeung > p Then p = pJeung

    If p = 0 Then Exit Function

    i = p + 1
    If i > Len(s) Then Exit Function

    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If IsBondRatingChar(ch) Then Exit Do
        i = i + 1
    Loop
    buf = ""
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If IsBondRatingChar(ch) Then
            buf = buf & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    BondRatingFromAsCell = Trim$(buf)
End Function

' needle 마지막 출현 위치(1-based), 없으면 0
Private Function LastInStr1(ByVal s As String, ByVal needle As String) As Long
    Dim p As Long
    Dim lastPos As Long
    lastPos = 0
    p = InStr(1, s, needle, vbTextCompare)
    Do While p > 0
        lastPos = p
        p = InStr(p + 1, s, needle, vbTextCompare)
    Loop
    LastInStr1 = lastPos
End Function

Private Function IsBondRatingChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    If ch >= "A" And ch <= "Z" Then IsBondRatingChar = True: Exit Function
    If ch >= "a" And ch <= "z" Then IsBondRatingChar = True: Exit Function
    If ch = "+" Or ch = "-" Then IsBondRatingChar = True: Exit Function
    If ch >= "0" And ch <= "9" Then IsBondRatingChar = True
End Function

Private Function FormatRate(ByVal v As Variant) As String
    If IsNumeric(v) Then
        FormatRate = CStr(CDbl(v))
    Else
        FormatRate = Trim$(CStr(v))
    End If
End Function

Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim m As Long, t As Long
    m = DATA_FIRST_ROW
    t = ws.Cells(ws.Rows.Count, COL_CP_ISSUER).End(xlUp).Row
    If t > m Then m = t
    t = ws.Cells(ws.Rows.Count, COL_BD_ISSUER).End(xlUp).Row
    If t > m Then m = t
    t = ws.Cells(ws.Rows.Count, ROW_META_FIRST_COL).End(xlUp).Row
    If t > m Then m = t
    GetLastDataRow = m
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        Nz = ""
    Else
        Nz = CStr(v)
    End If
End Function

Private Function CsvQuote(ByVal s As String) As String
    Dim t As String
    t = Replace(s, """", """""")
    If InStr(t, ",") > 0 Or InStr(t, Chr$(34)) > 0 Or InStr(t, vbCr) > 0 Or InStr(t, vbLf) > 0 Then
        CsvQuote = """" & t & """"
    Else
        CsvQuote = t
    End If
End Function

Private Function JoinCsvRow(ByVal fields As Variant) As String
    Dim i As Long
    Dim a() As String
    ReDim a(LBound(fields) To UBound(fields))
    For i = LBound(fields) To UBound(fields)
        a(i) = CsvQuote(CStr(fields(i)))
    Next i
    JoinCsvRow = Join(a, ",")
End Function

Private Sub WriteUtf8Csv(ByVal path As String, ByRef lines() As String, ByVal lineCount As Long)
    Dim stm As Object
    Dim i As Long
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Mode = 3
    stm.Open
    stm.Charset = "UTF-8"
    stm.WriteText ChrW(&HFEFF)
    For i = 1 To lineCount
        stm.WriteText lines(i) & vbLf
    Next i
    stm.SaveToFile path, 2
    stm.Close
End Sub
