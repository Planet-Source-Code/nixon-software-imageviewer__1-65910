Attribute VB_Name = "modMain"
Public Const vbTemporary = &H100
Public Const vbCompressed = &H800
Public Pan1 As Boolean

Public Function ErrorTrap()
If Err.Number = 0 Then Exit Function
Select Case Err.Number
Case 32755
Exit Function
Case 57
MsgBox "Device I/O error. Big Problems! Your hardware is acting up. Check the disc drive.", vbExclamation + vbOKOnly, "Error"
Case 61
MsgBox "The disk is full. Try deleting some files before saving.", vbExclamation + vbOKOnly, "Error"
Case 70
MsgBox "Access denied. Make sure the write protection is turned off at the source you are saving to.", vbExclamation + vbOKOnly, "Error"
Case 71
MsgBox "Device not ready. I think the drive door is open - please check.", vbExclamation + vbOKOnly, "Error"
Case 72
MsgBox "Disk media error. Time to inspect the media!", vbExclamation + vbOKOnly, "Error"
Case Else
MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, "Error"
End Select
End Function
