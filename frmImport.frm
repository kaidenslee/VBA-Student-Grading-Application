VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Import Database"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================================================+
' Date: April 3rd, 2025
' Program title: Import Database Userform
' Description: The interface importing a userform from anywhere in the computer
'===========================================================+

Private Sub UserForm_Initialize()
'instructions
    MsgBox "To use this tool, select the browse option and choose a file from your computer."
End Sub
    
'if user returns to main, hide this form and return to main userform
Private Sub btnBackToMain_Click()
'pass the database back to the main form
    frmGrades.txtFilePath.Text = Me.txtFilePath.Text
        Me.Hide
        frmGrades.Show
    End Sub

Private Sub btnBrowse_Click()
'code taken from lab on databases, open workbook path, allow user to select a file location, keep database
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.InitialFileName = ThisWorkbook.Path
    fd.Title = "Select registrar.mdb"
    fd.Filters.Clear
    fd.Filters.Add "Access Database", "*.mdb"
    
    If fd.Show = -1 Then
        txtFilePath.Text = fd.SelectedItems(1)
        MsgBox "Database selected!", vbInformation
    Else
        MsgBox "No file selected.", vbExclamation
    End If

    Set fd = Nothing
End Sub

