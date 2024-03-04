VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterForm 
   Caption         =   "Registration Page"
   ClientHeight    =   2412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5796
   OleObjectBlob   =   "RegisterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("Database")
    
    ' Fill database sheet from userform textbox values
    If TB1.Value <> "" Then
        For rowNum = 2 To 100
            If ws.Cells(rowNum, 1).Value = "" Then
                For cNum = 1 To 3
                    ws.Cells(rowNum, cNum) = Controls("TB" & cNum).Value
                Next cNum
                MsgBox "Registration Successful"
                Exit Sub
            End If
        Next rowNum
    Else
        MsgBox "Name cannot be empty."
    End If
End Sub
