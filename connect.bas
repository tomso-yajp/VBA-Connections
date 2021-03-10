Attribute VB_Name = "connect"
Const com As String = ",": Const at As String = "@"
Const adOpenKeyset As Integer = 1: Const adOpenDynamic As Integer = 2: Const adOpenStatic As Integer = 3
Const adLockReadOnly As Integer = 1: Const adLockPessimistic As Integer = 2: Const adLockOptimistic As Integer = 3
Const adCmdText As Integer = 1: Const adCmdTable As Integer = 2: Const adCmdStoredProc As Integer = 4

Sub debug_main()
Dim p As String, sh As String
sh = "sample"
p = ThisWorkbook.Path & "\" & sh & "."
'Create a sample file. file is an excel file and a text file
Call make_efile

'Output to the debug window.
'Connect to excel file using "ADODB.connection", "Microsoft.ACE.OLEDB.12.0"
Call output_debug(db_excel(p & "xlsm", sh), "db_excel  [Microsoft.ACE.OLEDB.12.0]")

'Output to the debug window.
'Connect to text file using "ADODB.connection", "Microsoft.ACE.OLEDB.12.0"
Call output_debug(db_text(p & "txt"), "db_text  [Microsoft.ACE.OLEDB.12.0]")

'Output to the debug window.
'Connect to excel file using "QueryTables.Add(ODBC,DSN,DBQ"
Call output_debug(qt_excel(p & "xlsm", sh), "qt_excel  [QueryTables.Add(ODBC,DSN,DBQ]")

'Output to the debug window.
'Connect to text file using "QueryTables.Add(ODBC,DSN,DBQ"
Call output_debug(qt_text(p & "txt"), "qt_text  [QueryTables.Add(Connection,Destination)]")

'Output to the debug window.
'Connect to text file using "ADODDB.Stream". encode is "shift_jis"
Call output_debug(st_text(p & "txt", "shift_jis"), "st_text   [ADODDB.Stream]")

'Output to the debug window.
'Connect to text file using "open path for as input #1"
Call output_debug(open_text(p & "txt", 0), "open_text   [open path for as input #1]")


'Output to the worksheet.
'Connect to excel file using "ADODB.connection", "Microsoft.ACE.OLEDB.12.0"
Call output_sheet(, db_excel(p & "xlsm", sh), "db_excel")

'Output to the worksheet.
'Connect to text file using "ADODB.connection", "Microsoft.ACE.OLEDB.12.0"
Call output_sheet(, db_text(p & "txt"), "db_text")

'Output to the worksheet.
'Connect to excel file using "QueryTables.Add(ODBC,DSN,DBQ"
Call output_sheet(, qt_excel(p & "xlsm", sh), "qt_excel")

'Output to the worksheet.
'Connect to text file using "QueryTables.Add(ODBC,DSN,DBQ"
Call output_sheet(, qt_text(p & "txt"), "qt_text")

'Output to the worksheet.
'Connect to text file using "ADODDB.Stream". encode is "shift_jis"
Call output_sheet(, st_text(p & "txt", "shift_jis"), "st_text")

'Output to the worksheet.
'Connect to text file using "open path for as input #1"
Call output_sheet(, open_text(p & "txt", 0), "open_text")

End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  qt_text : use Microsoft.ACE.OLEDB.12.0 to get the data in a excel file      XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function db_excel(Optional p As String, Optional sname As String)
Dim cn As Object, rs As Object
Dim i As Integer, ii As Integer
Dim t As Variant, tt As Variant
If p = "" Then p = Application.GetOpenFilename("Excel(*.xls*),*.xls*")
If p = "" Or Dir(p) = "" Then Exit Function
Set cn = CreateObject("adodb.connection")
With cn
  .Provider = "Microsoft.ACE.OLEDB.12.0"
  .Properties("Extended Properties") = "Excel 12.0 XML;HDR=NO;IMEX=1"
  .Open p
End With
Set rs = CreateObject("adodb.recordset")
rs.Open "select * from [" & sname & "$]", cn, adOpenStatic, adLockReadOnly
Do Until rs.EOF
  i = i + 1:
  For ii = 0 To rs.Fields.Count - 1
    tt = tt & rs.Fields(ii) & com
  Next
  t = t & Mid(tt, 1, Len(tt) - 1) & at
  tt = ""
  rs.MoveNext
Loop
rs.Close: cn.Close
Set cn = Nothing: Set rs = Nothing
db_excel = Mid(t, 1, Len(t) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  qt_text : use Microsoft.ACE.OLEDB.12.0 to get the data in a text file       XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function db_text(Optional p As String)
Dim cn As Object, rs As Object
Dim i As Integer, ii As Integer
Dim t As Variant, tt As Variant
Dim fname As String
If Dir(p) = "" Then Exit Function
With CreateObject("scripting.filesystemobject")
  If p = "" Then p = Application.GetOpenFilename("Ç∑Ç◊Çƒ(*.*),*.*")
  If p = "" Or Dir(p) = "" Then Exit Function
  fname = .GetFileName(p)
  p = .GetFile(p).ParentFolder
End With
Set cn = CreateObject("adodb.connection")
With cn
  .Provider = "Microsoft.ACE.OLEDB.12.0"
  .Properties("Data Source") = p & "\"
  .Properties("Extended Properties") = "text;HDR=NO;FMT=" & com & ";"
  .Open
End With
Set rs = CreateObject("adodb.recordset")
rs.Open "select * from [" & fname & "]", cn, adOpenStatic, adLockReadOnly, adCmdText
Do Until rs.EOF
  i = i + 1
  For ii = 0 To rs.Fields.Count - 1
    tt = tt & rs.Fields(ii) & com
  Next
  t = t & Mid(tt, 1, Len(tt) - 1) & at
  tt = ""
  rs.MoveNext
Loop
rs.Close: cn.Close
Set cn = Nothing: Set rs = Nothing
db_text = Mid(t, 1, Len(t) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  qt_text : use Microsoft.ACE.OLEDB.12.0 to get the data in a access file     XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub db_access()
Dim p As String, tname As String: p = "add path of *.accdb": tname = "add table name"
Dim cn As Object, rs As Object
Dim i As Integer, ii As Integer: i = 0: i = ii
Dim t As Variant, tt As Variant
Set cn = CreateObject("adodb.connection")
With cn
  .Provider = "Microsoft.ACE.OLEDB.12.0"
  .Properties("Data Source") = p
  .Open
End With
Set rs = CreateObject("adodb.recordset")
rs.Open "select * from etable", cn, adOpenStatic, adLockReadOnly
Do Until rs.EOF
  For ii = 0 To rs.Fields.Count - 1
    If i = 0 Then t = t & rs.Fields(ii).Name & com
    tt = tt & rs.Fields(ii) & com
  Next
  tt = tt & at
  rs.MoveNext: i = i + 1
Loop
t = t & at & Mid(tt, 1, Len(tt) - 1)
rs.Close: cn.Close
Set cn = Nothing
db_access = t
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  qt_text : get the data. use querytables.add to get the data in a excel file XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function qt_excel(Optional p As String, Optional sh As String = "make")
Dim constr As String, sql As String
If p = "" Then p = Application.GetOpenFilename("Excel(*.xls*),*.xls*")
If p = "" Or Dir(p) = "" Then Exit Function
constr = "ODBC;DSN=Excel Files;DBQ=" & p
sql = "select * from [" & sh & "$];"
With ThisWorkbook.Worksheets(add_sheet)
  .Cells.Clear
  With .QueryTables.Add(constr, .Cells(1, 1), sql)
    .SaveData = False
    .RefreshPeriod = 0
    .BackgroundQuery = False
    .RefreshStyle = xlOverwriteCells
    .Refresh
    .Delete
  End With
  qt_excel = qt_sub(.UsedRange.Value2)
  Application.DisplayAlerts = False
  .Delete
  Application.DisplayAlerts = True
End With
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  qt_text : get the data. use querytables.add to get the data in a text file  XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function qt_text(Optional p As String)
Dim constr As String: constr = "TEXT;" & p
If p = "" Then p = Application.GetOpenFilename("Ç∑Ç◊Çƒ(*.*),*.*")
If p = "" Or Dir(p) = "" Then Exit Function
With ThisWorkbook.Worksheets(add_sheet)
  .Cells.Clear
  With .QueryTables.Add(Connection:=constr, Destination:=.Cells(1, 1))
    .TextFilePlatform = 932 '932 shift_jis, 65001 utf-8, 1200 utf-16
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .SaveData = False
    .RefreshPeriod = 0
    .BackgroundQuery = False
    .RefreshStyle = xlOverwriteCells
    .Refresh
    .Delete
  End With
  qt_text = qt_sub(.UsedRange.Value2)
  Application.DisplayAlerts = False
  .Delete
  Application.DisplayAlerts = True
End With
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  qt_text : change the array the string text                                  XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function qt_sub(Optional v As Variant)
Dim r As Variant
Dim i As Integer, ii As Integer
For i = 1 To UBound(v, 1)
  For ii = 1 To UBound(v, 2): r = r & v(i, ii) & com: Next
  r = Mid(r, 1, Len(r) - 1) & at
Next
qt_sub = Mid(r, 1, Len(r) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  delete_qt : delete the connection                                           XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub delete_qt()
Dim t As Variant
With ThisWorkbook.ActiveSheet
  For Each t In .QueryTables: t.Delete:  Next
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  st_text : get the data. use the adodb stream                                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function st_text(Optional p As String, Optional en As String = "shift_jis")
Dim t As Variant, tt As Variant
If p = "" Then p = Application.GetOpenFilename("Ç∑Ç◊Çƒ(*.*),*.*")
If p = "" Or Dir(p) = "" Then Exit Function
With CreateObject("adodb.stream")
  .Open
  .Type = 2 'adTypeText
  .Charset = en
  .LoadFromFile p
  t = .ReadText
  .Close
End With
t = Replace(t, vbCr & vbLf, at) 'need to think
If Right(t, 1) = at Then t = Mid(t, 1, Len(t) - 1)
st_text = t
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  st_outtext : use the adodb stream                                           XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub st_outtext(ByVal p As String, Optional en As String = "shift_jis")
Dim t As Variant
With CreateObject("adodb.stream")
  .Open
  .Charset = en
  .LoadFromFile p & ".txt"
  t = .ReadText
  .Close
  
  .Open
  .Charset = en
  .WriteText t, 0 'adWriteChar
  .SaveToFile p & "_copy.txt", 2  'adSaveCreateOverWrite
  .Close
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  open_text : get the data. use the open statement                            XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function open_text(Optional p As String, Optional n As Long = 1)
Dim t As Variant, tt As String, i As Long
If p = "" Then p = Application.GetOpenFilename("Ç∑Ç◊Çƒ(*.*),*.*")
If p = "" Or Dir(p) = "" Then Exit Function
Open p For Input As #1
  Do Until EOF(1)
    i = i + 1: Line Input #1, t: tt = tt & t & at
    If i = n Then: open_text = t: Exit Do
  Loop
Close #1
If n = 0 Then open_text = Mid(tt, 1, Len(tt) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  change_text : convert excel file to text file                               XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function change_text(ByVal p As String)
Dim tname As String
Application.DisplayAlerts = False
With CreateObject("excel.application")
  .Visible = False
  With .Workbooks.Open(p)
    tname = .Path & "\" & Mid(.Name, 1, InStrRev(.Name, ".")) & "txt"
    If Dir(tname) <> "" Then Kill tname
    .SaveAs Filename:=tname, FileFormat:=xlCSV:    .Close
  End With
End With
Application.DisplayAlerts = True
chenge_text = tname
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  output_debug : output the list in the worksheet                             XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub output_sheet(Optional wb As Variant, Optional v As Variant, Optional sh As String = "make")
Dim t As Variant, tt As Variant
Dim i As Integer, ii As Integer
If IsError(wb) Then Set wb = ThisWorkbook
t = Split(v, at)
With wb
  If InStr(com & get_sheet(wb) & com, com & sh & com) = 0 Then _
    .Worksheets.Add.Name = sh
  With .Worksheets(sh)
    For i = 0 To UBound(t)
      tt = Split(t(i), com)
      For ii = 0 To UBound(tt)
        .Cells(i + 1, ii + 1) = tt(ii)
      Next
    Next
  End With
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  output_debug : output the list in the debug window                          XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub output_debug(ByVal v As Variant, Optional mstr As String = "")
Dim t As Variant, tt As Variant
Dim i As Integer, ii As Integer
t = Split(v, at)

Debug.Print "================================================================="
Debug.Print "     macro:" & mstr & "   output data"
Debug.Print "-----------------------------------------------------------------"
For i = 0 To UBound(t)
  Debug.Print t(i)
Next
Debug.Print "================================================================="

End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  make_tfile : create excel file of the sample                                XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub make_efile()
Dim wb As Variant
Dim v As Variant, p As String
p = ThisWorkbook.Path & "\sample."
If Dir(p) <> "" Then Kill p & "xlsm"
v = open_text(make_tfile, 0)
Application.ScreenUpdating = False
Set wb = Workbooks.Add
With wb
  If Dir(p & "xlsm") <> "" Then Kill p & "xlsm"
  .SaveAs p & "xlsm", 52
  Call output_sheet(wb, v, "sample")
  .Save
  .Close
End With
Application.ScreenUpdating = True
Set wb = Nothing
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  make_tfile : create text file of the sample                                 XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function make_tfile()
Dim cols As Variant: cols = "no,col1,col2,col3,col4,col5,memo"
Dim v As Variant: v = "tanaka,sato,ueda,imai,endo,naito,osaka"
Dim c As Variant, t As String, p As String
Dim n As Integer, i As Integer, ii As Integer
v = Split(v, com): p = ThisWorkbook.Path & "\sample.txt"
With CreateObject("scripting.filesystemobject")
  With .createtextfile(p)
    .writeline cols
    For i = 1 To 15
      c = c & i & com
      c = c & Chr(WorksheetFunction.RandBetween(65, 69)) & com
      c = c & WorksheetFunction.RandBetween(1, 9) * 1.5 & com
      n = WorksheetFunction.RandBetween(1, 100)
      If n Mod 2 = 0 Then t = "ÅZ" Else t = "Å~"
      c = c & t & com
      c = c & WorksheetFunction.RandBetween(0, 9) & com
      n = WorksheetFunction.RandBetween(LBound(v), UBound(v))
      c = c & v(n) & com
      c = c & "memo" & i
      .writeline c
      c = ""
    Next
  End With
End With
make_tfile = p
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  add_sheet : add the sheet in the workbook                                   XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function add_sheet(Optional wb As Variant, Optional sh As String = "make")
If IsError(wb) Then Set wb = ThisWorkbook
If InStr(com & get_sheet(wb) & com, com & sh & com) = 0 Then _
  wb.Worksheets.Add.Name = sh
add_sheet = sh
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  get_sheet : get the sheet list in the workbook                              XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function get_sheet(Optional wb As Variant)
Dim s As Variant, sh As Variant
If IsError(wb) Then Set wb = ThisWorkbook
With wb
  For Each s In .Worksheets
    sh = sh & s.Name & com
  Next
End With
get_sheet = Mid(sh, 1, Len(sh) - 1)
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX  get_sheet : delete the worksheet                                            XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub del_sheets(Optional sh As String = "Sheet1")
Dim s As Variant
With ThisWorkbook
  For Each s In .Worksheets
    Application.DisplayAlerts = False
    If InStr(com & sh & com, com & s.Name & com) = 0 Then s.Delete
    Application.DisplayAlerts = True
  Next
End With
End Sub
