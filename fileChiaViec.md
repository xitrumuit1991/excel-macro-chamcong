# File CHIA VIEC
```
Sub CHIAVIEC_Click()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim SheetName As String
Dim InputFile As Workbook
Dim OutputFile As Workbook
Set InputFile = ThisWorkbook

'Set OutputFile = Workbooks.Open("X:\088 HSC TTTM\XUAT KHAU\CHIA VIEC 2020\PHAN CONG CONG VIEC 2019.xlsx", Password:="012697", WriteResPassword:="012697")
Set OutputFile = Workbooks.Open("/Users/nguyen/Desktop/PHAN CONG CONG VIEC 2019.xlsx")
'Set OutputFile = Workbooks.Open("C:\Users\tanhd.ho\Desktop\New folder\PHAN CONG CONG VIEC 2019.xlsx", Password:="012697", WriteResPassword:="012697")
'Set OutputFile = Workbooks.Open("\\hcm00102669\Scan\PHUONG THANH\PHAN CONG CONG VIEC 2019.xlsx", password:="ABC", WriteResPassword:="ABC")


InputFile.Sheets("HOM NAY").Activate
InputFile.Sheets("HOM NAY").Range("A4:M2000").Select
Selection.ClearContents

InputFile.Sheets("CHIA VIEC").Activate
InputFile.Sheets("CHIA VIEC").Range("a4:p4").Select
InputFile.Sheets("CHIA VIEC").Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

InputFile.Sheets("HOM NAY").Activate
Sheets("HOM NAY").Range("a4").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

'check sheet CHIA VIEC in output file, if not found create one
Dim existsSheetChiaViec As Boolean
For I = 1 To OutputFile.Worksheets.Count
    If OutputFile.Worksheets(I).Name = "CHIA VIEC" Then
        existsSheetChiaViec = True
    End If
Next I

If Not existsSheetChiaViec Then
    OutputFile.Worksheets.Add.Name = "CHIA VIEC"
End If

''''
'OutputFile

OutputFile.Sheets("CHIA VIEC").Activate
OutputFile.Sheets("CHIA VIEC").Range("A4:p2000").Select
Selection.ClearContents

InputFile.Sheets("CHIA VIEC").Activate
InputFile.Sheets("CHIA VIEC").Range("a4:p4").Select
 
InputFile.Sheets("CHIA VIEC").Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
OutputFile.Sheets("CHIA VIEC").Activate
Sheets("CHIA VIEC").Range("a4").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

'check sheet THONG KE NGAY in output file, if not found create one
Dim existsSheetThongKeNgay As Boolean
For I = 1 To OutputFile.Worksheets.Count
    If OutputFile.Worksheets(I).Name = "THONG KE NGAY" Then
        existsSheetThongKeNgay = True
    End If
Next I

If Not existsSheetThongKeNgay Then
    OutputFile.Worksheets.Add.Name = "THONG KE NGAY"
End If

'PASTE THONG KE NGAY
InputFile.Sheets("THONG KE NGAY").Activate
InputFile.Sheets("THONG KE NGAY").Range("A1:AB300").Copy
OutputFile.Sheets("THONG KE NGAY").Activate
Sheets("THONG KE NGAY").Range("A1:AB300").PasteSpecial Paste:=xlPasteValuesAndNumberFormats



'check sheet DASHBOARD in output file, if not found create one
Dim existsSheetDashboard As Boolean
For I = 1 To OutputFile.Worksheets.Count
    If OutputFile.Worksheets(I).Name = "DASHBOARD" Then
        existsSheetDashboard = True
    End If
Next I

If Not existsSheetDashboard Then
    OutputFile.Worksheets.Add.Name = "DASHBOARD"
End If

'PASTE DASHBOARD
InputFile.Sheets("DASHBOARD").Activate
InputFile.Sheets("DASHBOARD").Range("A1:AB300").Copy
OutputFile.Sheets("DASHBOARD").Activate
Sheets("DASHBOARD").Range("A1:AB300").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

'PASTE VALUE
InputFile.Sheets("CHIA VIEC").Activate
InputFile.Sheets("CHIA VIEC").Range("A4:H800").Copy
InputFile.Sheets("CHIA VIEC").Activate
Sheets("CHIA VIEC").Range("A4:E800").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

OutputFile.Close SaveChanges:=True
ActiveWorkbook.Save
'Call s
InputFile.Sheets("CHIA VIEC").Activate
Range("A3").Select


Selection.End(xlDown).Select
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

```
