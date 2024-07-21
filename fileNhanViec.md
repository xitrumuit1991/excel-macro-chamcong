# File NHAN VIEC 
```
Sub Button1_Click()

'*** EDIT HERE for copy file
'FileCopy "X:\088 HSC TTTM\XUAT KHAU\CHIA VIEC 2020\PHAN CONG CONG VIEC 2019.xlsx", "d:\PHAN CONG CONG VIEC 2019.xlsx"
'FileCopy "X:\088 HSC TTTM\XUAT KHAU\CHIA VIEC 2020\PHAN CONG CONG VIEC 2019.xlsx", "d:\PHAN CONG CONG VIEC 2019.xlsx"

Dim InputFile As Workbook
Dim OutputFile As Workbook

'*** EDIT HERE
''Set InputFile = Workbooks.Open("C:\Users\tanhd.ho\Desktop\New folder\PHAN CONG CONG VIEC 2019.xlsx", Password:="012697", ReadOnly:=True)
'Set InputFile = Workbooks.Open("D:\PHAN CONG CONG VIEC 2019.xlsx", Password:="012697", ReadOnly:=True)
'***Set input file for Macos
Set InputFile = Workbooks.Open("/Users/nguyen/Desktop/PHAN CONG CONG VIEC 2019.xlsx", ReadOnly:=True)


Set OutputFile = ThisWorkbook
'Set lietke = Workbooks.Open("X:\088 HSC TTTM\XUAT KHAU\CHIA VIEC 2019\SL CHUNG TU\SL CHUNG TU_CNTTRANG.xlsx")

InputFile.Sheets("CHIA VIEC").Activate
InputFile.Sheets("CHIA VIEC").Range("A3:T1000").Copy
OutputFile.Sheets("HOM NAY").Activate
OutputFile.Sheets("HOM NAY").Range("A2:T1000").PasteSpecial Paste:=xlPasteValuesAndNumberFormats


OutputFile.Sheets("NHAN VIEC").Activate
OutputFile.Sheets("NHAN VIEC").Range("B4:N200").ClearContents



OutputFile.Sheets("HOM NAY").Range("$A$3:$M$1000").AutoFilter Field:=5, Criteria1:=OutputFile.Sheets("NHAN VIEC").Range("a1").Value

OutputFile.Sheets("HOM NAY").Range("A3:M1000").Copy
OutputFile.Sheets("NHAN VIEC").Activate
OutputFile.Sheets("NHAN VIEC").Range("B3:N3").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
OutputFile.Sheets("HOM NAY").ShowAllData

InputFile.Sheets("THONG KE NGAY").Activate
InputFile.Sheets("THONG KE NGAY").Range("A1:T100").Copy
OutputFile.Sheets("THONG KE NGAY").Activate
OutputFile.Sheets("THONG KE NGAY").Range("A1:T100").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

'
'Start
'nguyen update code copy sheet dashboard from inputfile to output file
'check sheet Dashboard, if not found => create
Dim existsSheetDashboard As Boolean
For i = 1 To OutputFile.Worksheets.Count
    If OutputFile.Worksheets(i).Name = "DASHBOARD" Then
        existsSheetDashboard = True
    End If
Next i
If Not existsSheetDashboard Then
    Worksheets.Add.Name = "DASHBOARD"
End If
'Copy sheet dashboard from inputfile to output file
InputFile.Sheets("DASHBOARD").Activate
InputFile.Sheets("DASHBOARD").Range("A1:T200").Copy
OutputFile.Sheets("DASHBOARD").Activate
OutputFile.Sheets("DASHBOARD").Range("A1:T200").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'END



ActiveWorkbook.Save

InputFile.Close savechanges:=False

OutputFile.Sheets("NHAN VIEC").Activate
Call INPUT1
Call CN_UPDATE

End Sub
```
