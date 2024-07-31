```javascript
Sub CheckAndMakeSureSheetNameExistedInRemote(RemoteFile, sheetName)
    ' Add code here
    Dim existsSheetSwiftChung As Boolean
    'Check sheet swift chung co chua, chua co thi tao moi
    For I = 1 To RemoteFile.Worksheets.Count
        If RemoteFile.Worksheets(I).Name = sheetName Then
            existsSheetSwiftChung = True
        End If
    Next I
    If Not existsSheetSwiftChung Then
        RemoteFile.Worksheets.Add.Name = sheetName
        RemoteFile.Save
    End If
End Sub
Function GetLastRowsBySheetName(fileWorkbook, sheetName) As Integer
    
    GetLastRowsBySheetName = fileWorkbook.Worksheets(sheetName).Cells(Cells.Rows.Count, "A").End(xlUp).Row
     
End Function
Function isAllowActionTongHopSwift() As Boolean
     'Dim isOutOfWorkingHour As Boolean
      CurrentTime = Time()
       If CurrentTime >= TimeValue("17:10:00") And CurrentTime <= TimeValue("23:59:00") Then
            'MsgBox "In working hour. " & CurrentTime
            isAllowActionTongHopSwift = True
      Else
            MsgBox "ERROR" & vbCrLf & "Vui long chi tong hop SWIFT sau 17:10:00 " & vbCrLf & "Thoat"
            isAllowActionTongHopSwift = False
    End If
End Function
Sub Collect_SWIFT_Cuoi_Ngay_Click()
        Dim SHEET_TONG_HOP_SWIFT As String: SHEET_TONG_HOP_SWIFT = "SHEET_TONG_HOP_SWIFT"
        Dim SHEET_SWIFT_CHUNG As String: SHEET_SWIFT_CHUNG = "SHEET_SWIFT_CHUNG"
        Dim IsClearSheetSwiftChungAfterCopy As Boolean: IsClearSheetSwiftChungAfterCopy = True
        Dim FlagEnableTimeCheck As Boolean: FlagEnableTimeCheck = False
        Dim RemoteFileSwiftChung As Workbook
        Dim RemoteFileTongHopSwift As Workbook
        'EDIT HERE
        'EDIT HERE
        'EDIT HERE
        PathFileRemoteSwiftChung = "/Users/nguyen/Desktop/remote/remoteFileInDiskX.xlsx" 'File nam o X: EDITABLE
        PathFileRemoteTongHopSwift = "/Users/nguyen/Desktop/remote/remoteFileTongHopSwift.xlsx" 'File nam o X: EDITABLE
        START_ROW_SHEET_CHUNG = 2 'Edit here
        
        'open file remote
        Set RemoteFileSwiftChung = Workbooks.Open(PathFileRemoteSwiftChung)
        Set RemoteFileTongHopSwift = Workbooks.Open(PathFileRemoteTongHopSwift)
        ThisWorkbook.Activate
        
        
        'CheckAndMakeSureSheetNameExistedInRemote
        Call CheckAndMakeSureSheetNameExistedInRemote(RemoteFileTongHopSwift, SHEET_TONG_HOP_SWIFT)
        
        'Confirm action
        ConfirmResult = MsgBox("Xac nhan tong hop Swift cuoi ngay ?", vbQuestion & vbYesNo)
        If ConfirmResult = vbNo Then
            MsgBox "Thoat"
            Exit Sub
        End If
           
        
        'check time allow action
        If FlagEnableTimeCheck = True Then
            IsAllowAction = isAllowActionTongHopSwift()
            If IsAllowAction = False Then
                ThisWorkbook.Activate
                Exit Sub
            End If
        End If
        
        
        
        'Start Copy process
        LastRowSheetChung = GetLastRowsBySheetName(RemoteFileSwiftChung, SHEET_SWIFT_CHUNG)
        LastRowSheetTongHop = GetLastRowsBySheetName(RemoteFileTongHopSwift, SHEET_TONG_HOP_SWIFT)
        'MsgBox "LastRowSheetChung=" & LastRowSheetChung & " LastRowSheetTongHop=" & LastRowSheetTongHop
        
        'Select Range Sheet Chung
        RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Activate
        RangeIndexSheetChung = "A" & START_ROW_SHEET_CHUNG & ":P" & LastRowSheetChung
        MsgBox "Select range copy SheetChung=" & RangeIndexSheetChung
        RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Range(RangeIndexSheetChung).Select
        Selection.Copy
        
        'Paste
        RemoteFileTongHopSwift.Sheets(SHEET_TONG_HOP_SWIFT).Activate
        PasteRangeIndex = "A" & (LastRowSheetTongHop + 1)
        'MsgBox "Copy to sheet tong hop range=" & PasteRangeIndex
        RemoteFileTongHopSwift.Sheets(SHEET_TONG_HOP_SWIFT).Range(PasteRangeIndex).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        'End copy  process
        
        
         'clear file swift chung sau khi copy thanh cong
        If IsClearSheetSwiftChungAfterCopy = True Then
           ClearRangeIndex = "A" & START_ROW_SHEET_CHUNG & ":P" & LastRowSheetChung
           resultYesNo = MsgBox("Copy thanh cong. Ban co muon xoa content cua Sheet Swift Chung khong ?" & vbCrLf & "Range:" & ClearRangeIndex, vbQuestion & vbYesNo)
           If resultYesNo = vbYes Then
                RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Activate
                ClearRangeIndex = "A" & START_ROW_SHEET_CHUNG & ":P" & LastRowSheetChung
                'MsgBox "Range se bi xoa: ClearRangeIndex SheetChung=" & ClearRangeIndex
                RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Range(ClearRangeIndex).Select
                Selection.ClearContents
                MsgBox "Xoa noi dung trong SHEET_SWIFT_CHUNG thanh cong. Range " & ClearRangeIndex
           End If
        End If
        
        'Save & exit
        RemoteFileSwiftChung.Save
        RemoteFileTongHopSwift.Save
        
        'set current active workbook  Activate
        ThisWorkbook.Activate
        
        MsgBox "Tong Hop Swift thanh cong. "

End Sub
Sub CHIAVIEC_Click()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim sheetName As String
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
