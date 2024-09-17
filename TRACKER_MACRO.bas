Attribute VB_Name = "TRACKER_MACRO"
Sub macro_PendingAddRow()
' Adds new row to the Pending Leave Data Table
    ActiveSheet.ListObjects("T_PENDING").ListRows.Add 1
End Sub

Sub macro_LeaveAddRow()
' Adds new row to the Leave Data Table
    ActiveSheet.ListObjects("T_LEAVE").ListRows.Add 1
End Sub

Sub macro_DeclinedAddRow()
' Adds new row to the Leave Data Table
    ActiveSheet.ListObjects("T_DECLINED").ListRows.Add 1
End Sub
Sub PreviousMonth()
Range("B2").Value = Month(DateValue("1 " & Range("B6").Value & " " & Range("B4").Value))
If ActiveSheet.Range("B2").Value = 1 Then
    Exit Sub
Else:
    Range("B2").Value = Range("B2").Value - 1
    Range("B6").Value = MonthName(Range("B2").Value)
End If
End Sub

Sub NextMonth()
Range("B2").Value = Month(DateValue("1 " & Range("B6").Value & " " & Range("B4").Value))
If ActiveSheet.Range("B2").Value = 12 Then
    Exit Sub
Else:
    Range("B2").Value = Range("B2").Value + 1
    Range("B6").Value = MonthName(Range("B2").Value)
End If
End Sub

Private Sub Auto_Open()
Dim edate As Date, mbox As Variant, myuser As String, wbuser As String

Application.ScreenUpdating = False
'CHANGE THE DATE
edate = DateSerial(2024, 1, 1)

If Date > edate Then
    MsgBox "Oops! Your access to this utility has been expired." & vbCrLf & _
            "Please ask the concern person to get the updated utility.", vbCritical, "Outdated/Expired Version"
    mbox = Application.InputBox("Pls input the password/code to continue...", "Password")
        If mbox <> "RDCWFM" Then ThisWorkbook.Close False
End If
End Sub

