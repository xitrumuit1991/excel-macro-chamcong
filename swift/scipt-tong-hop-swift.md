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
        Dim SHEET_2019 As String: SHEET_2019 = "2019"
        PasswordSheet2019 = "ctxk2"
        Dim SHEET_SWIFT_CHUNG As String: SHEET_SWIFT_CHUNG = "SHEET_SWIFT_CHUNG"
        Dim IsClearSheetSwiftChungAfterCopy As Boolean: IsClearSheetSwiftChungAfterCopy = True
        Dim FlagEnableTimeCheck As Boolean: FlagEnableTimeCheck = False
        Dim RemoteFileSwiftChung As Workbook
        Dim RemoteFileTongHopSwift As Workbook
        Dim LastRowSheet2019 As Long
        
        'EDIT HERE
        PathFileRemoteSwiftChung = "/Users/nguyen/Desktop/remote/remoteFileInDiskX.xlsx" 'File nam o X: EDITABLE
        START_ROW_SHEET_CHUNG = 2 'Edit here
        
        'open file remote
        Set RemoteFileSwiftChung = Workbooks.Open(PathFileRemoteSwiftChung)
        Set OutputFile2019 = ThisWorkbook
       
        OutputFile2019.Activate
        
        
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
                OutputFile2019.Activate
                Exit Sub
            End If
        End If
        
        'tim index row dau tien trong SHeet Chung co ngayGhi cotQ la hom nay (today) "yyyy_MM_dd" string
        LastRowSheetChung = GetLastRowsBySheetName(RemoteFileSwiftChung, SHEET_SWIFT_CHUNG)
        
        OutputFile2019.Sheets(SHEET_2019).Unprotect PasswordSheet2019 'unlock sheet
        LastRowSheet2019 = OutputFile2019.Sheets(SHEET_2019).UsedRange.Rows.Count + 1 ' +1 boi vi row 1 (empty) fucntion UsedRange se khong dem
        OutputFile2019.Sheets(SHEET_2019).Protect PasswordSheet2019 'lock sheet
        'MsgBox "LastRowSheet2019 = " & LastRowSheet2019
        
        IndexFirstRowToday = -1
        todayString = Format(Now, "yyyy") & "_" & Format(Now, "MM") & "_" & Format(Now, "dd")
        For idx = 2 To LastRowSheetChung
            rowNgayGhi = RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Range("Q" & idx).Value
            If todayString = rowNgayGhi Then
                IndexFirstRowToday = idx
                Exit For
            End If
        Next idx
        'MsgBox "LastRowSheetChung=" & LastRowSheetChung & " LastRowSheetTongHop=" & LastRowSheetTongHop
        If IndexFirstRowToday = -1 Or IndexFirstRowToday > LastRowSheetChung Then
            MsgBox "Khong co dong nao de copy"
            OutputFile2019.Sheets(SHEET_2019).Protect PasswordSheet2019 'lock sheet
            Exit Sub
        End If
        'MsgBox "IndexFirstRowToday = " & IndexFirstRowToday
        
        'Start Copy process
        'Copy cot A Sheet chung -> Cot A Sheet 2019
        'Select Range Sheet Chung
        RangeFromIndex = "A" & IndexFirstRowToday & ":A" & LastRowSheetChung 'A10:A20
        RangeToIndex = "A" & (LastRowSheet2019 + 1)
        OutputFile2019.Sheets(SHEET_2019).Unprotect PasswordSheet2019 'unlock sheet
        RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Range(RangeFromIndex).Copy
        OutputFile2019.Sheets(SHEET_2019).Range(RangeToIndex).PasteSpecial Paste:=xlPasteValues
        OutputFile2019.Sheets(SHEET_2019).Protect PasswordSheet2019 'lock sheet
        MsgBox "Copy Cot A Sheet Chung Range: (" & RangeFromIndex & ") sang cot A Sheet 2019 Range: (" & RangeToIndex & ") thanh cong"
        
        
         'Copy cot M Sheet chung -> Cot E Sheet 2019
        'Select Range Sheet Chung
        RangeFromIndex = "M" & IndexFirstRowToday & ":M" & LastRowSheetChung 'M10:M20
        RangeToIndex = "E" & (LastRowSheet2019 + 1)
        OutputFile2019.Sheets(SHEET_2019).Unprotect PasswordSheet2019 'unlock sheet
        RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Range(RangeFromIndex).Copy
        OutputFile2019.Sheets(SHEET_2019).Range(RangeToIndex).PasteSpecial Paste:=xlPasteValues
        OutputFile2019.Sheets(SHEET_2019).Protect PasswordSheet2019 'lock sheet
        MsgBox "Copy Cot M Sheet Chung Range: (" & RangeFromIndex & ") sang cot E Sheet 2019 Range: (" & RangeToIndex & ") thanh cong"
        
        
        
         'clear file swift chung sau khi copy thanh cong
        If IsClearSheetSwiftChungAfterCopy = True Then
           ClearRangeIndex = "A" & IndexFirstRowToday & ":Z" & LastRowSheetChung
           resultYesNo = MsgBox("Copy thanh cong. Ban co muon xoa content cua Sheet Swift Chung khong ?" & vbCrLf & "Range:" & ClearRangeIndex, vbQuestion & vbYesNo)
           If resultYesNo = vbYes Then
                RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Activate
                'ClearRangeIndex = "A" & IndexFirstRowToday & ":Z" & LastRowSheetChung
                'MsgBox "Range se bi xoa: ClearRangeIndex SheetChung=" & ClearRangeIndex
                RemoteFileSwiftChung.Sheets(SHEET_SWIFT_CHUNG).Range(ClearRangeIndex).Select
                Selection.ClearContents
                MsgBox "Xoa noi dung trong SHEET_SWIFT_CHUNG thanh cong. Range " & ClearRangeIndex
           End If
        End If
        
        'Save & exit
        RemoteFileSwiftChung.Save
        
        'set current active workbook  Activate
        'ThisWorkbook.Activate
        OutputFile2019.Sheets(SHEET_2019).Protect PasswordSheet2019 'lock sheet
        OutputFile2019.Save
        OutputFile2019.Activate
        
        MsgBox "Tong Hop Swift thanh cong. "

End Sub
```
