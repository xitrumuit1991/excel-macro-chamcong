```javascript
Function GetDateFromString(valueDateString, formatString)
    'Split space char
    arrDateTime = Split(valueDateString, " ")
    dateString = arrDateTime(0)
    timeString = arrDateTime(1)
    If formatString = "dd/mm/yy" Or formatString = "dd/mm/yyyy" Or formatString = "dd/mm/yy hh:mm:ss" Then
        arrDates = Split(dateString, "/")
        dd = arrDates(0)
        mm = arrDates(1)
        yyyy = arrDates(2)
        If Len(yyyy) = 2 Then
            yyyy = "20" & yyyy
        End If
        arrTimes = Split(timeString, ":")
        hhour = arrTimes(0)
        mminus = arrTimes(1)
        ssecond = arrTimes(2)
        
        fullDateString = yyyy & "-" & mm & "-" & dd & " " & hhour & ":" & mminus & ":" & ssecond
        
        GetDateFromString = CDate(fullDateString)
    Else
        GetDateFromString = Now
    End If
End Function
Function IsWorkBookOpen(FileName)
    Dim ff As Long, ErrNo As Long
    
     If Dir(FileName) <> vbNullString Then
        On Error Resume Next
        slot = FreeFile
        Open FileName For Binary Access Read Lock Read As #slot
        IsWorkBookOpen = Err.Number <> 0
        Close #slot
     End If
End Function
Function IsCorrectMaNhanVien(OutputFile, employeeCheck) As Boolean
    IsCorrectMaNhanVien = False
    If employeeCheck = "" Then
        MsgBox "Err employeeCheck is empty"
        Exit Function
    End If
    
    For i = 2 To 22
        maNhanVien = OutputFile.Worksheets("THONG KE NGAY").Range("B" & i).Value
        If maNhanVien <> "" And maNhanVien = employeeCheck Then
            IsCorrectMaNhanVien = True
            Exit For
        End If
    Next i
End Function
Sub CheckAndMakeSureSheetNameExistedInRemote(RemoteFile, sheetName)
    ' Add code here
    Dim existsSheetSwiftChung As Boolean
    'Check sheet swift chung co chua, chua co thi tao moi
    For i = 1 To RemoteFile.Worksheets.Count
        If RemoteFile.Worksheets(i).Name = sheetName Then
            existsSheetSwiftChung = True
        End If
    Next i
    If Not existsSheetSwiftChung Then
        RemoteFile.Worksheets.Add.Name = sheetName
        RemoteFile.Save
    End If
End Sub
Sub CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, sheetName)
    
    RemoteFile.Save
    RemoteFile.Close
    
    OutputFile.Save
    OutputFile.Worksheets(sheetName).Activate
    'RemoteFile.Close
    
End Sub
Sub CheckAndMakeSureSheetNameExistedInLocal(OutputFile, sheetName)
    Dim existsSheetSwiftCaNhan As Boolean
        'Check sheet swift ca nhanco chua, chua co thi tao moi
    For i = 1 To OutputFile.Worksheets.Count
        If OutputFile.Worksheets(i).Name = sheetName Then
            existsSheetSwiftCaNhan = True
        End If
    Next i
    If Not existsSheetSwiftCaNhan Then
        OutputFile.Worksheets.Add.Name = sheetName
        OutputFile.Save
    End If
End Sub
Sub UnlockRemoteFile(RemoteFile, SHEET_SWIFT_CHUNG)
     RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A1").Value = "" 'unlock file
    RemoteFile.Save
End Sub
Function isOutOfWorkingHour() As Boolean
     'Dim isOutOfWorkingHour As Boolean
      CurrentTime = Time()
       If CurrentTime >= TimeValue("08:00:00") And CurrentTime <= TimeValue("11:30:00") Then
            'MsgBox "In working hour. " & CurrentTime
            isOutOfWorkingHour = False
            
      ElseIf CurrentTime >= TimeValue("13:00:00") And CurrentTime <= TimeValue("17:10:00") Then
            'MsgBox "In working hour. " & CurrentTime
            isOutOfWorkingHour = False
            
      Else
            'MsgBox "ERROR" & vbCrLf & "Vui long get swift trong thoi gian lam viec 8:00->11:30 hoac 13:00->17:10" & vbCrLf & "Thoat"
            isOutOfWorkingHour = True
            
        End If
    'Return isOutOfWorkingHour
End Function
Function GetLastRowsBySheetName(fileWorkbook, sheetName) As Integer
    'LastRowSwiftCaNhan = OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Cells(Cells.Rows.Count, "A").End(xlUp).row
     'LastRowSwiftChung = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Cells(Cells.Rows.Count, "A").End(xlUp).row
    GetLastRowsBySheetName = fileWorkbook.Worksheets(sheetName).Cells(Cells.Rows.Count, "A").End(xlUp).row
     
End Function
Function GetLastColumnsBySheetName(fileWorkbook, sheetName, rowIndex) As Integer
    GetLastColumnsBySheetName = fileWorkbook.Worksheets(sheetName).Cells(rowIndex, Columns.Count).End(xlToLeft).Column
     
End Function

Function GetRowIndexByRefId(fileWorkbook, refIdString, sheetName) As Integer
    foundIdx = -1
    If refIdString = "" Or refIdString = vbNullString Then
        GetRowIndexByRefId = -1
        Exit Function
    End If
    
    LastRowSwiftCaNhan = fileWorkbook.Worksheets(sheetName).Cells(Cells.Rows.Count, "A").End(xlUp).row
    For idx = 2 To LastRowSwiftCaNhan
            refIdValue = fileWorkbook.Worksheets(sheetName).Range("A" & idx).Value
            If refIdValue <> "" And refIdString <> "" And refIdValue = refIdString Then
                foundIdx = idx
                'Exit For
            End If
            If foundIdx <> -1 Then Exit For
    Next idx
    GetRowIndexByRefId = foundIdx
     
End Function
Sub ButtonTraSwift_Click()
    Dim PathFileRemote As String

    PathFileRemote = "/Users/nguyen/Desktop/remote/remoteFileInDiskX.xlsx" 'File nam o X: EDITABLE
    'PathFileRemote = = "X:\088 HSC TTTM\XUAT KHAU\CHIA VIEC 2020/remoteFileInDiskX.xlsx" 'Path of VCB
    HardCodeEmployeeDung = "LQDUNG" 'Hardcode ten nhan vien cho the revert swift
     SHEET_SWIFT_CHUNG = "SHEET_SWIFT_CHUNG" 'ten sheet Chung
     SHEET_SWIFT_CA_NHAN = "SHEET_SWIFT_CA_NHAN" 'ten sheet Ca Nhan
    
    
    'Check file ton tai
    If Len(Dir(PathFileRemote)) = 0 Then
      MsgBox "File khong ton tai. " & PathFileRemote
      Exit Sub
    End If
    
    'Check file co available de  su dung  k ?
    IsWorkbookRemoteOpened = IsWorkBookOpen(PathFileRemote)
    If IsWorkbookRemoteOpened = True Then
        Set Wb = Workbooks.Open(FileName:=PathFileRemote, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
        If Wb.WriteReserved Then
          MsgBox "CODE_101. File remote sheet chung dang duoc su dung boi " & Wb.WriteReservedBy
          Wb.Close
          Exit Sub
        Else
          MsgBox "CODE_102. File remote sheet chung dang duoc su dung. Vui long doi"
          Wb.Close
          Exit Sub
        End If
    End If


     'Set file nhan viec
    Set RemoteFile = Workbooks.Open(PathFileRemote)
    Set OutputFile = ThisWorkbook
    
    'Check workbook readonly ? Xem co nguoi open truoc do chua ?. 'Check to see if file is already open
    If RemoteFile.ReadOnly Then
        MsgBox "CODE_103. File remote sheet chung dang duoc su dung. Vui long doi."
        RemoteFile.Close
        Exit Sub
    End If

    'step 1 check employee Name (k phai  Dung moi dc allow)
    'Check existed employeeName & set employeeName = Cell(a1)
    employeeName = OutputFile.Worksheets("NHAN VIEC").Range("A1").Value
    If Len(employeeName) = 0 Or employeeName = "" Or employeeName = vbNullString Then
        MsgBox "Chua nhap ten nhan vien"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    End If
    
    If UCase(employeeName) = UCase(HardCodeEmployeeDung) Then
        MsgBox "Chuc nang khong kha dung voi ban"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    End If
    
    
    'Step2:
    LastRowSheetChung = GetLastRowsBySheetName(RemoteFile, SHEET_SWIFT_CHUNG)
    LastRowSheetCaNhan = GetLastRowsBySheetName(OutputFile, SHEET_SWIFT_CA_NHAN)
    For idx = 2 To LastRowSheetChung
        transferFromValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("N" & idx).Value
        transferToValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("O" & idx).Value
        transferStatusValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("P" & idx).Value
        refIdValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A" & idx).Value
        
        '2.1 'Check existed  1 Transfer Swift roi =>check current employee phai la ng gui khong?
        'Neu current employee = nguoi gui (cot O), check status?
        If transferFromValue = employeeName Then 'Ban chinh la nguoi gui
            '2.1.1 Neu status = PENDING => bao loi
            If transferStatusValue = "PENDING" Then 'Trang thai dang pending (cho` ng nhan tra loi)
                'IsAlreadyExistTransferSwift = True
                MsgBox "Ban da gui di mot yeu cau transfer swift RefId (" & refIdValue & ") cho nguoi nhan (" & transferToValue & "). Hien tai trang thai la " & transferStatusValue & ". Vui long lien he nguoi nhan."
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
            End If
            
            '2.1.2 Neu status = APPROVE (nguoi nhan dong y) =>xoa refId trong sheetCaNhan & xoa 3 cot N, O, P (from, to ,status) trong sheetChung & update assigned nguoi nhan cho refId
            If transferStatusValue = "APPROVE" Then 'Trang thai  APPROVE (ng nhan accept yeu cau)
                
                MsgBox "Nguoi nhan " & transferToValue & " da Accept yeu cau cua ban. RefId (" & refIdValue & ")"
                
                'Xoa refId nay duoi sheet ca nhan.
                OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Rows(idx).Delete Shift:=xlUp
                MsgBox "Xoa refId o file local thanh cong. Row index bi xoa = " & idx
                
                
                'Edit cot status = 'TRANSFER_COMPLETE' =>de nguoi nhan change Name & xoa 3 cot N O P
                RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("N" & idx).Value = "" 'Xoa nguoi gui
                RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("O" & idx).Value = "" 'Xoa nguoi nhan
                RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("P" & idx).Value = "" 'Xoa status
                RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("M" & idx).Value = transferToValue 'Update assigned
                RemoteFile.Save
                OutputFile.Save
                MsgBox "Xoa RefId (" & refIdValue & ") tren file local thanh cong. Update thong tin Swift cho sheet Chung thanh cong. Thoat"
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
            End If 'End 2.1.2
            
            '2.1.3 Bao loi voi nhung status khong xac dinh (vd REJECT, ...)
            MsgBox "Dang co 1 Transfer Swift RefId (" & refIdValue & ") duoc gui tu ban voi trang thai " & transferStatusValue & ". Vui long kiem tra va xu ly"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            Exit Sub
            
        End If 'End 2.1
        
        '2.2 'Neu currentEmployee = nguoi nhan, check status? (ban la ng nhan transfer swift)
         If transferToValue = employeeName Then 'Ban chinh la nguoi nhan
            '2.2.1
            If transferStatusValue = "PENDING" Then 'Trang thai dang pending. Ban can phai confirm APPROVE/REJECT yeu cau
                'IsAlreadyExistTransferSwift = True
                confirmResult = MsgBox("Ban nhan dc mot yeu cau Transfer Swift RefId (" & refIdValue & ")  tu nguoi gui (" & transferFromValue & ") voi trang thai " & transferStatusValue & ". Ban vui long chon action: Yes (Accept) No (Reject)", vbQuestion + vbYesNo)
                If confirmResult = vbYes Then 'Approve request. Neu ACCEPT, copy 1 row tu SheetChung ve sheetCaNhan & update status = APPROVE
                    'Doi assign qua cho currentEmployee
                    RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("M" & idx).Value = employeeName
                    
                    'Copy row tu Sheet Chung => Tao 1 row trong sheet CaNhan
                    RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Activate
                    RangeCopied = "A" & (LastRowSheetCaNhan + 1) & ":" & "Z" & (LastRowSheetCaNhan + 1) 'Vd A1:Z1
                    RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Rows(idx).Copy _
                    Destination:=OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Range(RangeCopied)
                    
                    'update status = APPROVE
                    RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("P" & idx).Value = "APPROVE"
                    
                    RemoteFile.Save
                    OutputFile.Save
                    MsgBox "Copy thong tin Swift RefId (" & refIdValue & ")  ve file local thanh cong, update status thanh APPROVE thanh cong."
                    Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                    Exit Sub
                Else 'Reject request
                    RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("P" & idx).Value = "REJECT"
                    MsgBox "Ban da REJECT yeu cau Transfer Swift RefId (" & refIdValue & ") . Vui long lien he nguoi gui: " & transferFromValue
                    Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                    Exit Sub
                End If
                
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
            End If 'End 2.2.1
            
            '2.2.2
            If transferStatusValue = "APPROVE" Then
                MsgBox "Ban da APPROVE yeu cau Transfer Swift RefId (" & refIdValue & ") nay roi. Vui long lien he nguoi gui: " & transferFromValue
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
            End If 'End 2.2.2
            
            '2.2.3
            If transferStatusValue = "REJECT" Then
                MsgBox "Ban da REJECT yeu cau Transfer Swift RefId (" & refIdValue & ") nay roi. Vui long lien he nguoi gui: " & transferFromValue
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
            End If 'End 2.2.3
            
            '2.2.4 Bao loi voi nhung status khong xac dinh
            MsgBox "Dang co 1 Transfer Swift RefId (" & refIdValue & ") duoc gui cho ban voi trang thai " & transferStatusValue & ". Vui long kiem tra va xu ly"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            Exit Sub
            
         End If 'End 2.2
        
    Next idx
    
    
    LastRowSheetChung = GetLastRowsBySheetName(RemoteFile, SHEET_SWIFT_CHUNG)
    LastRowSheetCaNhan = GetLastRowsBySheetName(OutputFile, SHEET_SWIFT_CA_NHAN)
    'Step3:
    '3.1 currentEmployee k phai nguoi gui/nguoi nhan (k co 1 Transfer Swift nao da Tao truoc day)
    '=> Show popup cho phep input refId & ten nguoi nhan
    Dim nguoiNhan As String
    nguoiNhan = InputBox("Ban muon Tra Swift cho ai ? Vui long nhap ten nguoi nhan: ") 'Show popup cho phep input refId & ten nguoi nhan
    If nguoiNhan <> "" Then
      'MsgBox "Nguoi nhan la " & nguoiNhan
      isCorrectName = IsCorrectMaNhanVien(OutputFile, nguoiNhan)
      If isCorrectName = False Then
            MsgBox "Ten nguoi nhan khong dung. Kiem tra lai theo Range B2:B22 trng Sheet THONG KE NGAY. Thoat"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            Exit Sub
      End If
    Else
        MsgBox "Ban chua nhap ten nguoi nhan. Thoat"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    End If
    
    '3.2 Show popup cho phep input refId
    inputRefId = InputBox("Vui long nhap Ref Id muon tra: ")
    If inputRefId <> "" Then
      'MsgBox "Input RefId = " & inputRefId
    Else
        MsgBox "Ban chua nhap RefId cua Swift. Thoat"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    End If
    
    foundRefInCaNhan = GetRowIndexByRefId(OutputFile, inputRefId, SHEET_SWIFT_CA_NHAN)
    foundRefInChung = GetRowIndexByRefId(RemoteFile, inputRefId, SHEET_SWIFT_CHUNG)
    'Check RefId co hop le (nam trong SheetCaNhan va nam trong Sheet Chung) => k hop le bao loi
    ''MsgBox "foundRefInCaNhan=" & foundRefInCaNhan
    ''MsgBox "foundRefInChung=" & foundRefInChung
    If foundRefInCaNhan = -1 Or foundRefInChung = -1 Then
        MsgBox "Ref Id (" & inputRefId & ") khong co trong Sheet SWIFT_CANHAN hoac khong co trong Sheet SWIFT_CHUNG !!! " & vbCrLf & " Vui long kiem tra lai lai RefId"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    End If
    
     '3.4 Update value cac cot transferFromValue N, O, P (from, to ,status)
     RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("N" & foundRefInChung).Value = employeeName
     RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("O" & foundRefInChung).Value = nguoiNhan
     RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("P" & foundRefInChung).Value = "PENDING"
     RemoteFile.Save
     OutputFile.Save
     MsgBox "Transfer Swift thanh cong." & vbCrLf & " Update Swift RefId (" & inputRefId & ")  trong sheet Chung thanh cong. Nguoi gui " & employeeName & ". Nguoi nhan " & nguoiNhan & ". Trang thai PENDING"
     Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
     Exit Sub
    
End Sub
Sub ButtonGetSwift_Click()
'khai bao hang` const
Dim START_ROW_SHEET_INPUT As Integer
Dim START_ROW_SHEET_CHUNG As Integer
Dim NUMBER_SWIFT_CAN_GET As Integer
Dim SHEET_SWIFT_CHUNG As String
Dim SHEET_SWIFT_CA_NHAN As String

Dim FlagRevertSwift As Boolean
Dim FlagCheckOutOfWorkHour As Boolean
Dim FlagLimitSwiftCanGet As Boolean
Dim HardCodeEmployeeDung As String
Dim FlagEnableTransferSwift As String: FlagEnableTransferSwift = True
Dim IsWorkbookRemoteOpened As Boolean
Dim isLockedRemoteFile As String


'EDIT HERE
'EDIT HERE
Dim PathFileRemote As String
PathFileRemote = "/Users/nguyen/Desktop/remote/remoteFileInDiskX.xlsx" 'File nam o X: EDITABLE
'PathFileRemote = = "X:\088 HSC TTTM\XUAT KHAU\CHIA VIEC 2020/remoteFileInDiskX.xlsx" 'Path of VCB

 START_ROW_SHEET_INPUT = 22
 START_ROW_SHEET_CHUNG = 2
 NUMBER_SWIFT_CAN_GET = 3 'limit swift co the Get ve
 HardCodeEmployeeDung = "LQDUNG" 'Hardcode ten nhan vien cho the revert swift

 SHEET_SWIFT_CHUNG = "SHEET_SWIFT_CHUNG" 'ten sheet Chung
 SHEET_SWIFT_CA_NHAN = "SHEET_SWIFT_CA_NHAN" 'ten sheet Ca Nhan
 FlagRevertSwift = True 'ON/OFF tinh nang revert Swift cho nhan vien
 FlagCheckOutOfWorkHour = True 'ON/OFF tinh nang ngoai gio lam viec
 FlagLimitSwiftCanGet = True 'ON/OFF tinh nang limit so luon swift lay ve


'Special case for LQDUNG
Dim IsRevertSwiftAction As Integer
Dim RevertSwiftRefId As String


'khai bao bien
Dim RemoteFile As Workbook
Dim OutputFile As Workbook 'File nam local
Dim employeeName As String 'ten nhan vien Sheet INPUT
Dim SelectedRefSwift As String
Dim SelectedIndexSwift As Integer
'Set ngay ghi = today
todayString = Format(Now, "yyyy") & "_" & Format(Now, "MM") & "_" & Format(Now, "dd")

'Check file ton tai
If Len(Dir(PathFileRemote)) = 0 Then
  MsgBox "File khong ton tai. " & PathFileRemote
  Exit Sub
End If

'Check file co available de  su dung  k ?
IsWorkbookRemoteOpened = IsWorkBookOpen(PathFileRemote)
If IsWorkbookRemoteOpened = True Then
    Set Wb = Workbooks.Open(FileName:=PathFileRemote, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    If Wb.WriteReserved Then
      'Debug.Print "File in use by " & Wb.WriteReservedBy
      MsgBox "CODE_101. File remote sheet chung dang duoc su dung boi " & Wb.WriteReservedBy
      Wb.Close
      Exit Sub
    Else
      'Debug.Print "File in use by another one"
      MsgBox "CODE_102. File remote sheet chung dang duoc su dung. Vui long doi"
      Wb.Close
      Exit Sub
    End If
End If

'Set file nhan viec
Set RemoteFile = Workbooks.Open(PathFileRemote)
Set OutputFile = ThisWorkbook

'Check workbook readonly ? Xem co nguoi open truoc do chua?. 'Check to see if file is already open
If RemoteFile.ReadOnly Then
    MsgBox "CODE_103. File remote sheet chung dang duoc su dung. Vui long doi."
    RemoteFile.Close
    Exit Sub
End If

'Check sheet swift chung co chua, chua co thi tao moi
Call CheckAndMakeSureSheetNameExistedInRemote(RemoteFile, SHEET_SWIFT_CHUNG)

'Check sheet swift ca nhan Local co chua?, chua co thi tao moi
Call CheckAndMakeSureSheetNameExistedInLocal(OutputFile, SHEET_SWIFT_CA_NHAN)

'Check sheet swift SHEET_SWIFT_CHUNG o local co chua, chua co thi tao moi
Call CheckAndMakeSureSheetNameExistedInLocal(OutputFile, SHEET_SWIFT_CHUNG)


'Check existed employeeName & set employeeName = Cell(a1)
employeeName = OutputFile.Worksheets("NHAN VIEC").Range("A1").Value
If Len(employeeName) = 0 Or employeeName = "" Or employeeName = vbNullString Then
    MsgBox "Chua nhap ten nhan vien"
    OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Activate 'Active sheet
    Exit Sub
End If


'check out of working hour
FlagCheckOutOfWorkHour = False
If FlagCheckOutOfWorkHour = True Then
    Dim isOutOfWorkingValue As Boolean: isOutOfWorkingValue = isOutOfWorkingHour()
      If isOutOfWorkingValue = True Then
            MsgBox "ERROR" & vbCrLf & "Vui long get swift trong thoi gian lam viec 8:00->11:30 hoac 13:00->17:10" & vbCrLf & "Thoat"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            
            Exit Sub
        End If
End If

'Transfer Swift cho dong nghiep khac, Bao Error khi co 1 transfer chua xu ly
If FlagEnableTransferSwift = True And UCase(employeeName) <> UCase(HardCodeEmployeeDung) Then
    'Check ban dang co 1 yeu cau transfer swift tu dong nghiep.
    LastRowSheetChung = GetLastRowsBySheetName(RemoteFile, SHEET_SWIFT_CHUNG)
    LastRowSheetCaNhan = GetLastRowsBySheetName(OutputFile, SHEET_SWIFT_CA_NHAN)
    For idx = 2 To LastRowSheetChung
        transferFromValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("N" & idx).Value
        transferToValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("O" & idx).Value
        transferStatusValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("P" & idx).Value
        refIdValue = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A" & idx).Value
        
        If transferFromValue = employeeName Then
                MsgBox "Ban da gui di mot yeu cau Transfer Swift RefId=(" & refIdValue & ") cho (" & transferToValue & "). Vui long xy ly."
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
        End If
        
         If transferToValue = employeeName Then
                MsgBox "Ban da nhan duoc mot yeu cau Transfer Swift RefId=(" & refIdValue & ") gui tu (" & transferFromValue & "). Vui long xy ly."
                Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                Exit Sub
        End If
    Next idx
End If


'Empployee DUNG  Revert SWIFT: Neu la Empployee DUNG click button => show confirm ACTION (get/revert Swift)
If UCase(employeeName) = UCase(HardCodeEmployeeDung) Then
    IsRevertSwiftAction = MsgBox("Ban co muon tra lai Swift nao khong ?" & vbCrLf & "Yes (Tra lai)." & vbCrLf & "No (Nhan Swift moi)", vbQuestion + vbYesNo)
    If IsRevertSwiftAction = vbYes Then
        RevertSwiftRefId = InputBox("Vui long nhap Swift RefId de tra lai.")
        If RevertSwiftRefId = "" Then
            MsgBox "Ban chua nhap RefId. Thoat"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            Exit Sub
        End If
        
        'Check swift RefId co nam trong sheet ca nhan khong?
        IdxFoundInSwiftCaNhan = -1
        IdxFoundInSwiftChung = -1
        LastRowSwiftCaNhan = OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Cells(Cells.Rows.Count, "A").End(xlUp).row
        LastRowSwiftChung = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Cells(Cells.Rows.Count, "A").End(xlUp).row
        For idx = 2 To LastRowSwiftCaNhan
            RefIdCaNhan = OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Range("A" & idx).Value
            If RefIdCaNhan = RevertSwiftRefId Then
                IdxFoundInSwiftCaNhan = idx
            End If
        Next idx
        
        For idx = 2 To LastRowSwiftChung
            RefIdChung = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A" & idx).Value
            If RefIdChung = RevertSwiftRefId Then
                IdxFoundInSwiftChung = idx
            End If
        Next idx
        
        If IdxFoundInSwiftCaNhan = -1 Then
            MsgBox "RefId tra lai (" & RevertSwiftRefId & ") khong ton tai trong SHEET_SWIFT_CA_NHAN." & vbCrLf & "Thoat"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            Exit Sub
        End If
        If IdxFoundInSwiftChung = -1 Then
            MsgBox "RefId tra lai (" & RevertSwiftRefId & ") khong ton tai trong SHEET_SWIFT_CHUNG." & vbCrLf & "Thoat"
            Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
            Exit Sub
        End If
        
        OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Rows(IdxFoundInSwiftCaNhan).Delete Shift:=xlUp
        RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Rows(IdxFoundInSwiftChung).Delete Shift:=xlUp
        
        OutputFile.Save
        RemoteFile.Save
        
        MsgBox "Tra lai Swift RefId=(" & RevertSwiftRefId & ") thanh cong. " & vbCrLf & _
            "Row index sheet ca nhan bi xoa =" & IdxFoundInSwiftCaNhan & vbCrLf & _
            "Row index sheet chung bi xoa = " & IdxFoundInSwiftChung & vbCrLf & _
            "Thoat"
        
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    Else
        'MsgBox "Tiep tuc nhan Swift..."
    End If
End If



'Check limit number swift can get per day NUMBER_SWIFT_CAN_GET
If FlagLimitSwiftCanGet = True Then
    CountNumberSwift = 0
    LastRowInput = OutputFile.Worksheets("INPUT").Cells(Cells.Rows.Count, "A").End(xlUp).row
    LastRowSwiftCaNhan = OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Cells(Cells.Rows.Count, "A").End(xlUp).row
    LastRowSwiftChung = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Cells(Cells.Rows.Count, "A").End(xlUp).row
    
    ''Loop sheet SWIFT CHUNG
    For idx = 2 To LastRowSwiftChung
        RefIdCheck = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A" & idx).Value 'cot A
        AssignerIdCheck = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("M" & idx).Value 'cot M
        ngayGhi = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("Q" & idx).Value 'todayString
        
         'row nay dang assign current employee & ngayGhi = today
        If AssignerIdCheck = employeeName And ngayGhi = todayString Then
                 'Loop sheet INPUT
                For idy = 2 To LastRowInput
                    RefIdInINPUT = OutputFile.Worksheets("INPUT").Range("A" & idy).Value
                    StatusRefId = OutputFile.Worksheets("INPUT").Range("J" & idy).Value
                    If RefIdCheck = RefIdInINPUT And StatusRefId = "Input" Then
                        CountNumberSwift = CountNumberSwift + 1
                    End If
                Next idy
        End If
    Next idx
    
    Dim maxSwiftCanGet As Integer: maxSwiftCanGet = NUMBER_SWIFT_CAN_GET
    If employeeName = HardCodeEmployeeDung Then
        maxSwiftCanGet = 10 'Set lai max swift dc phep nhan cua employee
    End If
    
    If CountNumberSwift >= maxSwiftCanGet Then
        MsgBox "So luong Ref Swift (status Input) cua ban vuot qua gioi han.  " & CountNumberSwift & " >= " & maxSwiftCanGet & vbCrLf & " Ban khong the nhan them. " & vbCrLf & " Thoat"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub
    End If
End If
    
'check file remote co dang bi locked k ? Neu co thi bao loi~
OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Activate
countdownTimeCheck = 3
MsgBox "Xin chao ban:  " & employeeName & vbCrLf & "Click de bat dau kiem tra trong " & countdownTimeCheck & " giay ..."
For idx = 1 To countdownTimeCheck
    Application.StatusBar = "Kiem tra tinh kha dung cua Sheet SHEET_SWIFT_CHUNG trong " & (countdownTimeCheck - idx) & " giay"
    
    isLockedRemoteFile = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A1").Value
    If isLockedRemoteFile = "locked" Then
        MsgBox "File remote dang co nguoi su dung (A1=locked). Vui long thu lai sau"
        Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
        Exit Sub 'Thoat khoi chuong trinh
    End If
   
    Application.Wait (Now + TimeValue("00:00:01"))  'Sleep 1000
Next idx
Application.StatusBar = "Sheet SHEET_SWIFT_CHUNG san sang de su dung"

'File chua ai su dung, set locked
RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A1").Value = "locked"
RemoteFile.Save

OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Activate 'active lai sheet ca nhan





''Tinh last row va last column cua tung sheet
NumRowsInput = GetLastRowsBySheetName(OutputFile, "INPUT") 'OutputFile.Worksheets("INPUT").Cells(Cells.Rows.Count, "A").End(xlUp).row
NumColumsInput = GetLastColumnsBySheetName(OutputFile, "INPUT", START_ROW_SHEET_INPUT) 'OutputFile.Worksheets("INPUT").Cells(START_ROW_SHEET_INPUT, Columns.Count).End(xlToLeft).Column
NumRowsSheetChung = GetLastRowsBySheetName(RemoteFile, SHEET_SWIFT_CHUNG)
NumColumsSheetChung = GetLastColumnsBySheetName(RemoteFile, SHEET_SWIFT_CHUNG, START_ROW_SHEET_CHUNG)

'MsgBox "Sheet INPUT: So dong= " & NumRowsInput & " So Cot=" & NumColumsInput & vbCrLf & " Sheet SHEET CHUNG: So dong= " & NumRowsSheetChung & " So Cot=" & NumColumsSheetChung

'START
'START
'START Loop tung row trong sheet INPUT,
SelectedRefSwift = ""
SelectedIndexSwift = -1
For i = 2 To NumRowsInput
    Dim isPickedLC As Boolean
    Dim nguoiNhan As String
    isPickedLC = False 'Khai bao var
    cotLValue = OutputFile.Worksheets("INPUT").Cells(i, "L").Value 'Cot L phai la SWIFT
    cotJValue = OutputFile.Worksheets("INPUT").Cells(i, "J").Value 'Cot J value phai = "Input"
    isNACotLValue = Application.WorksheetFunction.IsNA(cotLValue)
    isNACotJValue = Application.WorksheetFunction.IsNA(cotJValue)
    
    'DK de get dc SWIFT: co 3 dk
    'DK1: cotLValue=SWIFT (sheet INPUT)
    'DK2: cotJValue = "Input" (sheet INPUT)
    'DK3: thoi gian nho nhat (thoi gian tinh ca ngay va gio, bao gom ca ngay qua khu)
    
    'Check DK1 va DK2: Chi~ lay row cotLValue=SWIFT & cotJValue = "Input"
    If isNACotLValue = False And isNACotJValue = False And cotJValue <> "#N/A" And cotLValue <> "#N/A" _
        And cotLValue = "SWIFT" And UCase(cotJValue) = "INPUT" Then
      
      'Lay Ref value
      RefValue = OutputFile.Worksheets("INPUT").Cells(i, "A").Value
      'MsgBox "dong " & i & " la row SWIFT." & " Ref Id = " & RefValue
      
             'START Loop tung row cua Sheet chung
             For j = 2 To NumRowsSheetChung
               RefVauleSheetChung = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Cells(j, "A").Value
               
               If RefValue = RefVauleSheetChung Then
                   'MsgBox "Info: Ref Id =" & RefValue & "  da co nguoi nhan. Nguoi nhan =" & NguoiNhan
                   isPickedLC = True
                   nguoiNhan = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Cells(j, "M").Value
                   Exit For 'Ref co ng nhan roi, thoat khoi loop for
               End If
               
             Next j
            'END Loop tung row cua Sheet chung
     
        'Neu Ref cua row nay` chua co nguoi nao nhan^ thi`
        If isPickedLC = False And cotLValue = "SWIFT" And UCase(cotJValue) = "INPUT" Then
           
           If SelectedIndexSwift = -1 Then
           '* Neu SelectedSwift empty => SelectedSwift = row nay.
                  SelectedRefSwift = RefValue
                  SelectedIndexSwift = i
                  'MsgBox "INFO: Tim thay FIRST Swift refId =(" & SelectedRefSwift & ") And index row= " & SelectedIndexSwift
           ElseIf SelectedIndexSwift > -1 Then
               '   * Neu da co' SelectedSwift truoc do => check thoi gian lay cai tre nhat
               
               currentSwiftTimeCotK = OutputFile.Worksheets("INPUT").Cells(SelectedIndexSwift, "K").Value
               newSwiftTimeCotK = OutputFile.Worksheets("INPUT").Cells(i, "K").Value
               
               If newSwiftTimeCotK = "" Or IsDate(newSwiftTimeCotK) = False Then
                    MsgBox "Value " & newSwiftTimeCotK & " tai range K" & i & " khong phai format Date"
                    Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
                    Exit Sub
               End If
               
               newSwiftDate = GetDateFromString(newSwiftTimeCotK, "dd/mm/yy") 'Macro doc value string k dung format => can custom function split string
               currentSwiftDate = GetDateFromString(currentSwiftTimeCotK, "dd/mm/yy")
               
               'MsgBox "INFO: Tim thay OTHER Swift refId =(" & RefValue & ") And index row= " & i & vbCrLf & _
               '     "Month current = " & Month(currentSwiftDate) & " Day current =" & Day(currentSwiftDate) & vbCrLf & _
               '      "Month new = " & Month(newSwiftDate) & " Day new =" & Day(newSwiftDate) & vbCrLf & _
               '      "DateDiff= " & DateDiff("s", newSwiftDate, currentSwiftDate)
                     
                'LC nao co thoi gian nho nhat  thi chon
               If DateDiff("s", newSwiftDate, currentSwiftDate) > 0 Then 'the same newSwiftTimeCotK < currentSwiftTimeCotK
                        'MsgBox "Ngay newSwiftDate < Ngay currentSwiftDate => chon RefId nay =" & RefValue & ". Row index= " & i
                       SelectedRefSwift = RefValue
                       SelectedIndexSwift = i
               End If
           End If
           
        End If 'End check isPickedLC
     
      
    End If 'End Check cotLValue = "SWIFT" And UCase(cotJValue) = "INPUT"
Next i

If SelectedIndexSwift = -1 Then
    Call UnlockRemoteFile(RemoteFile, SHEET_SWIFT_CHUNG)
    OutputFile.Save
    Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)
    MsgBox "INFO: Khong tim thay SWIFT nao hop le." & vbCrLf & "Thoat."
    Exit Sub
End If

'Cuoi cung lay dc cai row Swift chua co nguoi nhan & co thoi gan tre nhat
'MsgBox "INFO: Tim thay SWIFT : RefId =(" & SelectedRefSwift & "), tai row index = " & SelectedIndexSwift

'Copy RefSwift from INPUT vao Sheet SHEET_SWIFT_CA_NHAN & Assign Ten Nguoi Nhan
LastRowSheetCaNhan = OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Cells(Cells.Rows.Count, "A").End(xlUp).row + 1
fromCopyRange = "A" & SelectedIndexSwift & ":Q" & SelectedIndexSwift
toCopyRange = "A" & LastRowSheetCaNhan & ":Q" & LastRowSheetCaNhan 'Vd A10:L10

OutputFile.Worksheets("INPUT").Range(fromCopyRange).Copy _
Destination:=OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Range(toCopyRange)
OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Cells(LastRowSheetCaNhan, "M").Value = employeeName
OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Cells(LastRowSheetCaNhan, "Q").Value = todayString
OutputFile.Save
'MsgBox "INFO: added new swift into swift local success. From INPUT range " & fromCopyRange & " to SWIFT CA NHAN range " & toCopyRange


'Update/Push RefId Swift from SHEET_SWIFT_CA_NHAN vao Sheet SHEET_SWIFT_CHUNG.
LastRowSheetChung = RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Cells(Cells.Rows.Count, "A").End(xlUp).row + 1
RangeForCopiedSheetChung = "A" & LastRowSheetChung & ":" & "M" & LastRowSheetChung 'Vd A10:M10
OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Rows(LastRowSheetCaNhan).Copy _
Destination:=RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range(RangeForCopiedSheetChung)

RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("Q" & LastRowSheetChung).Value = todayString
RemoteFile.Save
'MsgBox "INFO: added new swift into remote file success"

'unlock file
Call UnlockRemoteFile(RemoteFile, SHEET_SWIFT_CHUNG) 'RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A1").Value = "" 'RemoteFile.Save


'Copy Sheet SHEET_SWIFT_CHUNG ve local
RemoteFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A1:T500").Copy
OutputFile.Worksheets(SHEET_SWIFT_CHUNG).Activate
OutputFile.Worksheets(SHEET_SWIFT_CHUNG).Range("A1:T500").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'MsgBox "INFO: Copy sheet swift chung ve local thanh cong"


'Get swift xong
MsgBox "INFO: Get swift thanh cong." & vbCrLf & "SWIFT Ref Id = " & SelectedRefSwift '& " tai row index = " & SelectedIndexSwift

'OutputFile.Activate
'OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Activate

RemoteFile.Save 'Close SaveChanges:=False
RemoteFile.Close SaveChanges:=False

OutputFile.Save
OutputFile.Worksheets(SHEET_SWIFT_CA_NHAN).Activate
'Call CloseRemoteFileAndActiveSheetCaNhan(RemoteFile, OutputFile, SHEET_SWIFT_CA_NHAN)

End Sub
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

InputFile.Close SaveChanges:=False

OutputFile.Sheets("NHAN VIEC").Activate
Call INPUT1
Call CN_UPDATE

End Sub
Function GetTimeRemainSLA(typeCheck As String)
    Dim i As Integer
    Dim COLUMN_MAPPING_LOAI_GD As String, COLUMN_MAPPING_LOAI_GD_VALUE As String
    Dim COLUMN_START_ROW As Integer
    
    COLUMN_MAPPING_LOAI_GD = "Y"
    COLUMN_MAPPING_LOAI_GD_VALUE = "Z"
    COLUMN_START_ROW = 2
    
    For i = COLUMN_START_ROW To Sheets("NHAN VIEC").Range(COLUMN_MAPPING_LOAI_GD & Rows.Count).End(xlUp).row + 1 Step 1
        Dim loaiGD As String
        loaiGD = Sheets("NHAN VIEC").Range(COLUMN_MAPPING_LOAI_GD & i).Value
        If loaiGD <> "" And loaiGD = typeCheck Then
            GetTimeRemainSLA = Sheets("NHAN VIEC").Range(COLUMN_MAPPING_LOAI_GD_VALUE & i).Value
            Exit For
        End If
    Next i
End Function
Function IsTimeSANG(dateCheck As String)
    If Format(dateCheck, "hh:mm:ss") >= "11:31:00" And Format(dateCheck, "hh:mm:ss") <= "17:00:00" Then
        IsTimeSANG = False
    Else
        IsTimeSANG = True
    End If
End Function
Function IsTimeCHIEU(dateCheck As String)
    If Format(dateCheck, "hh:mm:ss") >= "11:31:00" And Format(dateCheck, "hh:mm:ss") <= "17:00:00" Then
        IsTimeCHIEU = True
    Else
        IsTimeCHIEU = False
    End If
End Function

Sub Button3_Click()
    'Dim loaiGiaoDich As String, loaiKhachHang As String, gioChiNhanhUp As String
    Dim totalRows As Integer, i As Integer, j As Integer
    Dim wsSLA As Worksheet, wsBangMa As Worksheet
    Dim wbSource As Workbook
    
    'config cot
    Dim START_ROW As String, COLUMN_GIOCN_UP As String, COLUMN_LOAI_GD As String, COLUMN_LOAI_KH As String, COLUMN_REF As String
    Dim COLUMN_GIO_UP_RESET As String, COLUMN_DEADLINE As String, COLUMN_SO_PHUT As String
    START_ROW = 4
    COLUMN_GIOCN_UP = "X"
    COLUMN_LOAI_GD = "P"
    COLUMN_LOAI_KH = "G"
    COLUMN_REF = "B"
    COLUMN_GIO_UP_RESET = "U"
    COLUMN_DEADLINE = "W"
    COLUMN_SO_PHUT = "V"
    COLUMN_SO_LUONG_CHUNG_TU = "O" 'so luong chung tu
    
    Set wsSLA = Sheets("NHAN VIEC") 'set worksheet
    totalRows = wsSLA.Range("A" & Rows.Count).End(xlUp).row 'lastRow of SLA
    
    wsSLA.Range(COLUMN_GIO_UP_RESET & START_ROW & ":" & COLUMN_GIO_UP_RESET & totalRows).Clear    'Clear Column E
    wsSLA.Range(COLUMN_SO_PHUT & START_ROW & ":" & COLUMN_SO_PHUT & totalRows).Clear
    wsSLA.Range(COLUMN_DEADLINE & START_ROW & ":" & COLUMN_DEADLINE & totalRows).Clear
   
    Dim ii As Integer
    
    'Format date GioCN Up;  'Format Data Column D (dd/mm/yyyy hh:mm:ss)
    For ii = START_ROW To totalRows Step 1
        Dim subsDate() As String
        Dim tmpCNUp As String
        
        If wsSLA.Cells(ii, COLUMN_GIOCN_UP) <> "" Then
            subsDate = Split(Cells(ii, COLUMN_GIOCN_UP), "/")
            If subsDate(0) <> "" And subsDate(1) <> "" And subsDate(2) <> "" And Len(subsDate(0)) = 2 And Len(subsDate(1)) = 2 Then
                Dim subsYear() As String
                subsYear = Split(subsDate(2))
                If subsYear(0) <> "" And subsYear(0) <> "2020" And Len(subsYear(0)) = 2 And subsYear(1) <> "" Then
                    tmpCNUp = subsDate(0) & "/" & subsDate(1) & "/" & "2020"
                    If tmpCNUp <> "" Then
                         tmpCNUp = tmpCNUp & " " & subsYear(1)
                         Cells(ii, COLUMN_GIOCN_UP).NumberFormat = "dd/MM/yyyy hh:mm:ss"
                         Cells(ii, COLUMN_GIOCN_UP) = tmpCNUp
                         'MsgBox tmpCNUp
                     End If
                End If
            End If
        End If
    Next ii
    
    'Start loop each row & generate deadline
    For i = START_ROW To totalRows Step 1
        Dim rowRef As String
        Dim rowLoaiGD As String
        Dim rowLoaiKH As String
        Dim rowGioCNUp As String
        Dim timeRemain As Integer
        Dim soluongChungTu As Integer
        
        rowRef = wsSLA.Cells(i, COLUMN_REF)
        rowLoaiGD = wsSLA.Cells(i, COLUMN_LOAI_GD) 'loai giao dich column O
        rowLoaiKH = wsSLA.Cells(i, COLUMN_LOAI_KH) 'loai KH column G
        rowGioCNUp = wsSLA.Cells(i, COLUMN_GIOCN_UP) 'gio chi nhanh Up column X
        soluongChungTu = wsSLA.Cells(i, COLUMN_SO_LUONG_CHUNG_TU)
        
        'Neu  Gio Up thuoc GioNghiTrua thi=> set lai thanh 13h
        If Format(rowGioCNUp, "hh:mm:ss") >= "11:30:00" And Format(rowGioCNUp, "hh:mm:ss") < "13:00:00" Then
            If Format(rowGioCNUp, "hh:mm:ss") > "11:30:59" Then
                rowGioCNUp = Format(rowGioCNUp, "yyyy-mm-dd") & " 13:00:00"
            Else
                rowGioCNUp = Format(rowGioCNUp, "yyyy-mm-dd") & " 11:30:00"
            End If
            Cells(i, COLUMN_GIO_UP_RESET).Value = rowGioCNUp
            Cells(i, COLUMN_GIO_UP_RESET).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End If
        
        
        
        'Neu GioUp > 17h => set lai thanh  17h
        If Format(rowGioCNUp, "hh:mm:ss") >= "17:00:00" Then
            rowGioCNUp = Format(rowGioCNUp, "yyyy-mm-dd") & " 17:00:00"
            rowGioCNUp = DateAdd("n", (15 * 60), rowGioCNUp)
            Cells(i, COLUMN_GIO_UP_RESET).Value = rowGioCNUp
            Cells(i, COLUMN_GIO_UP_RESET).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End If
        
        'Neu GioUp >0h & GioUp <8h Sang => set lai 8h Sang
        If Format(rowGioCNUp, "hh:mm:ss") > "00:00:00" And Format(rowGioCNUp, "hh:mm:ss") < "08:00:00" Then
            rowGioCNUp = Format(rowGioCNUp, "yyyy-mm-dd") & " 08:00:00"
            Cells(i, COLUMN_GIO_UP_RESET).Value = rowGioCNUp
            Cells(i, COLUMN_GIO_UP_RESET).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End If
        
        
        If rowRef <> "" And rowLoaiGD <> "" And rowGioCNUp <> "" Then
            Dim keyCheckTimeRemain As String
            Dim IsSang As Boolean, IsChieu As Boolean
            Dim deadLine As String
            
            IsSang = IsTimeSANG(rowGioCNUp) 'Check time is Sang or chieu
            IsChieu = IsTimeCHIEU(rowGioCNUp)

            keyCheckTimeRemain = rowLoaiGD & rowLoaiKH 'Prepair Key To Map with Sheet 2
            timeRemain = GetTimeRemainSLA(keyCheckTimeRemain)
            
            
            'Note: he so so luong chung tu
            'If rowLoaiKH = "BLUEDIA" Or rowLoaiKH = "DIA" Or rowLoaiKH = "GOL" Then
                'If rowLoaiGD = "OC-ORG" Or rowLoaiGD = "RV1-ORG" Or rowLoaiGD = "RV1" Or rowLoaiGD = "RV2" Or rowLoaiGD = "RV2-ORG" Or rowLoaiGD = "RV3" Or rowLoaiGD = "RV4" Then
                    'If soluongChungTu >= 0 And soluongChungTu <= 6 Then
                        'timeRemain = GetTimeRemainSLA(keyCheckTimeRemain)
                    'ElseIf soluongChungTu > 6 Then
                        'timeRemain = GetTimeRemainSLA(keyCheckTimeRemain) * soluongChungTu / 6
                    'End If
                'End If
            'End If
            
            
            'Khong tinh sang & chieu
            'If timeRemain = 0 Then 'If timeRemain = 0 add SANG  OR CHIEU into Key => check again
            '    If IsSang = True Then
            '        keyCheckTimeRemain = keyCheckTimeRemain & "SANG"
            '    Else
            '        keyCheckTimeRemain = keyCheckTimeRemain & "CHIEU"
            '    End If
            '    timeRemain = GetTimeRemainSLA(keyCheckTimeRemain)
            'End If
            
            
            
            If timeRemain > 0 Then
                deadLine = DateAdd("n", timeRemain, rowGioCNUp)
                'MsgBox "deadLine" & deadLine
                
                'TH1: Neu gio UP va Deadline < Gio Nghi Trua 11h30
                If Format(rowGioCNUp, "hh:mm:ss") <= "11:30:00" And Format(deadLine, "hh:mm:ss") <= "11:30:00" Then
                    Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                    Cells(i, COLUMN_DEADLINE).Value = deadLine
                    Cells(i, COLUMN_SO_PHUT).Value = timeRemain 'Show number minus plus
                
                'TH2 Neu Gio ChiNhanhUp < GioNghiTrua & Deadline > GioNghiTrua
                ElseIf Format(rowGioCNUp, "hh:mm:ss") <= "11:30:00" And Format(deadLine, "hh:mm:ss") > "11:30:00" Then
                    deadLine = DateAdd("n", (timeRemain + 90), rowGioCNUp) 'plus 90phut nghi trua
                    If Format(deadLine, "hh:mm:ss") > "17:00:00" Then
                        deadLine = DateAdd("n", (timeRemain + 90 + 15 * 60), rowGioCNUp)
                        Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                        Cells(i, COLUMN_DEADLINE).Value = deadLine
                        Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & "90 + 15*60 (nghi trua + qua ngay)" 'Show number minus plus
                        If Format(deadLine, "hh:mm:ss") > "11:30:00" Then
                            deadLine = DateAdd("n", (timeRemain + 90 + 15 * 60 + 90), rowGioCNUp)
                            Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                            Cells(i, COLUMN_DEADLINE).Value = deadLine
                            Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & "90 + 15*60 + 90 (trua + qua ngay + trua)" 'Show number minus plus
                        End If
                    Else
                        Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                        Cells(i, COLUMN_DEADLINE).Value = deadLine
                        Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & "90 (nghi trua)" 'Show number minus plus
                    End If
                    
                'TH3 Neu GioChiNhanhUp >= 13h & Deadline <= 17h => giu nguyen deadline
                ElseIf Format(rowGioCNUp, "hh:mm:ss") >= "11:30:00" And Format(rowGioCNUp, "hh:mm:ss") <= "17:00:00" And Format(deadLine, "hh:mm:ss") >= "11:30:00" And Format(deadLine, "hh:mm:ss") <= "17:00:00" Then
                    Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                    Cells(i, COLUMN_DEADLINE).Value = deadLine
                    Cells(i, COLUMN_SO_PHUT).Value = timeRemain 'Show number minus plus
                    
                'TH4 Neu GioUp > 13h & GioUp <= 17h & Deadline > 17h => deadline + 15h  (cong them 15tieng qua sang mai)
                ElseIf Format(rowGioCNUp, "hh:mm:ss") >= "11:30:00" And Format(rowGioCNUp, "hh:mm:ss") <= "17:00:00" And Format(deadLine, "hh:mm:ss") > "17:00:00" And Format(deadLine, "hh:mm:ss") < "23:59:00" Then
                    Dim minusPlus As Integer
                    minusPlus = 15 * 60  ' plus 15*60phut qua ngay hom sau
                    deadLine = DateAdd("n", (timeRemain + minusPlus), rowGioCNUp)
                    If Format(deadLine, "hh:mm:ss") > "11:30:00" Then
                        deadLine = DateAdd("n", (timeRemain + minusPlus + 90), rowGioCNUp) 'Plus More 90 Phut  if >Time Nghi Trua (ngay hom sau)
                        Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                        Cells(i, COLUMN_DEADLINE).Value = deadLine
                        Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & (minusPlus + 90) & " (trua hom sau)" 'Show number minus plus
                    Else
                        Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                        Cells(i, COLUMN_DEADLINE).Value = deadLine
                        Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & minusPlus & " (Sang hom sau)"  'Show number minus plus
                    End If
                    
                'TH5 GioCN  Up trong  (13h, 15h) va ngay Dealine # ngay CN Up
                ElseIf Format(rowGioCNUp, "hh:mm:ss") > "13:00:00" And Format(rowGioCNUp, "hh:mm:ss") <= "17:00:00" Then
                        If Format(deadLine, "yyyy-mm-dd") > Format(rowGioCNUp, "yyyy-mm-dd") Then
                            deadLine = DateAdd("n", (timeRemain + 15 * 60), rowGioCNUp)
                            If Format(deadLine, "hh:mm:ss") > "11:30:00" Then
                                deadLine = DateAdd("n", (timeRemain + 15 * 60 + 90), rowGioCNUp)
                                Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                Cells(i, COLUMN_DEADLINE).Value = deadLine
                                Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & 15 * 60 & " +90 (hom sau + trua)" 'Show number minus plus
                            Else
                                Cells(i, COLUMN_DEADLINE).NumberFormat = "dd/mm/yyyy hh:mm:ss"
                                Cells(i, COLUMN_DEADLINE).Value = deadLine
                                Cells(i, COLUMN_SO_PHUT).Value = timeRemain & " + " & 15 * 60 & " (hom sau)" 'Show number minus plus
                            End If
                            
                        End If
                End If
                
            End If
            
            ''''Show
            'MsgBox "Ref = " & rowRef & "; keyCheckTimeRemain=" & keyCheckTimeRemain & " timeremain=" & timeRemain & "; rowGioCNUp=" & rowGioCNUp & "; IsSang=" & IsSang & "; IsChieu=" & IsChieu
            'MsgBox "deadLine=" & deadLine
            
        End If
    Next i
    
    ActiveWorkbook.Save
    MsgBox "DONE " & totalRows & " rows"
End Sub

```
