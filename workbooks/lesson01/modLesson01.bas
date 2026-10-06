Attribute VB_Name = "modLesson01"
Option Explicit

' Decoding VBA — Lesson 1: Tour the Visual Basic Editor
' https://decodingvba.com/learn/beginner/beginner-vbe-tour/
'

Sub HelloDecodingVBA()
    MsgBox "VBE is open. You are ready to write code.", vbInformation, "Decoding VBA"
End Sub

Sub ShowActiveSheetName()
    MsgBox "Active sheet: " & ActiveSheet.Name, vbInformation, "Decoding VBA"
End Sub
