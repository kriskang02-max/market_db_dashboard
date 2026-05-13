Attribute VB_Name = "BondDb_ExportToCsv"
'===============================================================================
' bond_db.xlsx → CSV (Overview 등에서 표 형태로 쓰기 쉬운 형식)
'
' 시트 구조 (동일 시트 내 두 블록):
'   국고채 — 헤더 D10:M10, 데이터 D11:M400
'   통안채 — 헤더 O10:X10, 데이터 O11:X100
'
' 출력 (워크북 저장 폴더 또는 고정 폴더):
'   bond_db_ktb.csv   국고채 (첫 행 = 헤더)
'   bond_db_msb.csv   통안채 (첫 행 = 헤더)
'
' 빈 행: 블록 하단에서 연속 빈 행은 출력에서 제외 (마지막 유효 행까지).
' 인코딩: UTF-8 BOM (Excel에서 한글 열기 호환)
'
' 실행:
'   1) bond_db.xlsx 열어둔 상태에서 ExportBondDbToCsv
'   2) 또는 ExportBondDbFromDefaultPath (경로 고정 오픈 후 내보내기)
'===============================================================================

Option Explicit

' 비워두면 첫 번째 시트, 지정하면 해당 이름 시트
Private Const BOND_SHEET_NAME As String = ""

Private Const HEADER_ROW As Long = 10

' 국고채 D:M (= 4~13)
Private Const KTB_COL_FIRST As Long = 4
Private Const KTB_COL_LAST As Long = 13
Private Const KTB_DATA_ROW_FIRST As Long = 11
Private Const KTB_DATA_ROW_LAST As Long = 400

' 통안채 O:X (= 15~24)
Private Const MSB_COL_FIRST As Long = 15
Private Const MSB_COL_LAST As Long = 24
Private Const MSB_DATA_ROW_FIRST As Long = 11
Private Const MSB_DATA_ROW_LAST As Long = 100

Private Const OUT_FILE_KTB As String = "bond_db_ktb.csv"
Private Const OUT_FILE_MSB As String = "bond_db_msb.csv"

Private Const DEFAULT_BOND_XLSX As String = "C:\Users\infomax\Documents\market_db_dashboard\bond_db.xlsx"

' Value2가 더블일 때 날짜로만 취급할 엑셀 일련번호 범위 (그 밖은 숫자 그대로)
Private Const XL_DATE_SERIAL_MIN As Double = 38000#  ' 약 2003년 부근
Private Const XL_DATE_SERIAL_MAX As Double = 55000#  ' 약 2040년 부근

'-------------------------------------------------------------------------------

Public Sub ExportBondDbToCsv()
    On Error GoTo EH
    RunExport ActiveWorkbook
    Exit Sub
EH:
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

Public Sub ExportBondDbFromDefaultPath()
    Dim wb As Workbook

    If Dir(DEFAULT_BOND_XLSX) = "" Then
        MsgBox "파일 없음: " & DEFAULT_BOND_XLSX, vbExclamation
        Exit Sub
    End If

    Set wb = Nothing
    On Error GoTo FailOpen
    Set wb = Workbooks.Open(DEFAULT_BOND_XLSX, ReadOnly:=True)
    wb.Activate
    RunExport wb
    GoTo CloseWb

FailOpen:
    MsgBox "열 수 없음: " & DEFAULT_BOND_XLSX & vbCrLf & Err.Description, vbCritical

CloseWb:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub RunExport(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim outDir As String

    Set ws = ResolveBondSheet(wb)
    If ws Is Nothing Then
        MsgBox "bond 시트를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    outDir = ResolveOutputDir(wb)
    If Len(outDir) = 0 Then
        outDir = Left$(DEFAULT_BOND_XLSX, InStrRev(DEFAULT_BOND_XLSX, "\") - 1)
        If Len(outDir) = 0 Then outDir = CurDir$
    End If

    Application.ScreenUpdating = False
    On Error GoTo EH

    ExportBlock ws, HEADER_ROW, KTB_COL_FIRST, KTB_COL_LAST, KTB_DATA_ROW_FIRST, KTB_DATA_ROW_LAST, _
                outDir & Application.PathSeparator & OUT_FILE_KTB, "국고채"

    ExportBlock ws, HEADER_ROW, MSB_COL_FIRST, MSB_COL_LAST, MSB_DATA_ROW_FIRST, MSB_DATA_ROW_LAST, _
                outDir & Application.PathSeparator & OUT_FILE_MSB, "통안채"

CleanExit:
    Application.ScreenUpdating = True
    MsgBox "저장 완료:" & vbCrLf & outDir & Application.PathSeparator & OUT_FILE_KTB & vbCrLf & _
           outDir & Application.PathSeparator & OUT_FILE_MSB, vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description, vbCritical
End Sub

Private Function ResolveBondSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    If Len(BOND_SHEET_NAME) > 0 Then
        Set ResolveBondSheet = wb.Worksheets(BOND_SHEET_NAME)
    Else
        Set ResolveBondSheet = wb.Worksheets(1)
    End If
End Function

Private Function ResolveOutputDir(ByVal wb As Workbook) As String
    Dim p As String
    p = wb.Path
    If Len(p) > 0 Then
        ResolveOutputDir = p
    Else
        ResolveOutputDir = ""
    End If
End Function

Private Sub ExportBlock(ByVal ws As Worksheet, _
                        ByVal headerRow As Long, _
                        ByVal colFirst As Long, ByVal colLast As Long, _
                        ByVal rowFirst As Long, ByVal rowLast As Long, _
                        ByVal outPath As String, _
                        ByVal blockLabel As String)

    Dim nCols As Long
    Dim lastDataRow As Long
    Dim hdr() As String
    Dim lines() As String
    Dim cap As Long
    Dim n As Long
    Dim r As Long, c As Long
    Dim ci As Long
    Dim fields() As String

    nCols = colLast - colFirst + 1
    If nCols < 1 Then Err.Raise vbObjectError + 1, , "열 범위 오류"

    ReDim hdr(1 To nCols)
    For c = colFirst To colLast
        hdr(c - colFirst + 1) = CellToCsvField(ws.Cells(headerRow, c).Value2)
    Next c

    lastDataRow = LastNonEmptyRowInBlock(ws, rowFirst, rowLast, colFirst, colLast)
    If lastDataRow < rowFirst Then
        Err.Raise vbObjectError + 2, , blockLabel & ": 데이터 행이 없습니다."
    End If

    cap = (lastDataRow - rowFirst + 1) + 4
    ReDim lines(1 To cap)
    n = 1
    lines(n) = JoinCsvRowStrings(hdr)

    For r = rowFirst To lastDataRow
        ReDim fields(1 To nCols)
        For ci = 1 To nCols
            fields(ci) = CellToCsvField(ws.Cells(r, colFirst + ci - 1).Value2)
        Next ci
        n = n + 1
        If n > UBound(lines) Then
            cap = cap * 2
            ReDim Preserve lines(1 To cap)
        End If
        lines(n) = JoinCsvRowStrings(fields)
    Next r

    WriteUtf8Csv outPath, lines, n
End Sub

Private Function LastNonEmptyRowInBlock(ByVal ws As Worksheet, _
                                       ByVal rowFirst As Long, ByVal rowLast As Long, _
                                       ByVal colFirst As Long, ByVal colLast As Long) As Long
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    For r = rowLast To rowFirst Step -1
        For c = colFirst To colLast
            v = ws.Cells(r, c).Value2
            If Not IsEmpty(v) Then
                If Len(Trim$(CStr(v))) > 0 Then
                    LastNonEmptyRowInBlock = r
                    Exit Function
                End If
            End If
        Next c
    Next r
    LastNonEmptyRowInBlock = rowFirst - 1
End Function

Private Function CellToCsvField(ByVal v As Variant) As String
    Dim d As Double
    On Error Resume Next
    If IsEmpty(v) Then
        CellToCsvField = ""
        Exit Function
    End If
    If VarType(v) = vbDate Then
        CellToCsvField = Format$(CDate(v), "yyyy-mm-dd")
        Exit Function
    End If
    If IsNumeric(v) Then
        d = CDbl(v)
        If d >= XL_DATE_SERIAL_MIN And d <= XL_DATE_SERIAL_MAX Then
            CellToCsvField = Format$(CDate(v), "yyyy-mm-dd")
            If Err.Number <> 0 Then
                Err.Clear
                CellToCsvField = Trim$(Replace(CStr(d), ",", ""))
            End If
            Exit Function
        End If
        CellToCsvField = Trim$(Replace(CStr(d), ",", ""))
    Else
        CellToCsvField = Trim$(CStr(v))
    End If
End Function

Private Function JoinCsvRowStrings(ByRef fields() As String) As String
    Dim i As Long
    Dim a() As String
    ReDim a(LBound(fields) To UBound(fields))
    For i = LBound(fields) To UBound(fields)
        a(i) = CsvQuote(fields(i))
    Next i
    JoinCsvRowStrings = Join(a, ",")
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
