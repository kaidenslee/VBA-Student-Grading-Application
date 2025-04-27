VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrades 
   Caption         =   "Student Grades"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmGrades.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ========================================================+
' Date: April 3rd, 2025
' Program title: Grades Userform (main)
' Description: The main interface for the grading application
'===========================================================+
Dim dbPath As String
Dim conn As Object, rs As Object
Dim sql As String
Dim finalGrades() As Double
Dim total As Double, count As Long, i As Long
Dim avg As Double, minG As Double, maxG As Double, stdev As Double
Dim sumSqDiff As Double
' if user presses cancel button, form will close
Private Sub cmdCancel_Click()
    Unload Me
End Sub
'if user selects display average button
Private Sub cmdDisplayAverage_Click()
'check database has been loaded first
If txtFilePath.Text = "" Then
        MsgBox "Please import the database first!", vbExclamation
        Exit Sub
    End If
    'hide current form and display the appropriate form
    Me.Hide
    frmClassAverage.Show

End Sub

'if user selects generate chart button
Private Sub cmdGenerateChart_Click()
'check database has been loaded first
If txtFilePath.Text = "" Then
        MsgBox "Please import the database first!", vbExclamation
        Exit Sub
    End If
    'hide current form and display the appropriate form
    Me.Hide
    frmChart.Show
End Sub


Private Sub cmdGenerateReport_Click()
' secure connection to database imported and ensure it has been loaded before continuing
    dbPath = frmGrades.txtFilePath.Text
    If dbPath = "" Then
        MsgBox "Please import the database first!", vbExclamation
        Exit Sub
    End If

    ' set connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath

    ' calculate final grades using weighted formula
    sql = "SELECT (A1*0.05 + A2*0.05 + A3*0.05 + A4*0.05 + MidTerm*0.3 + [Final Exam]*0.5) AS FinalGrade FROM Grades"
    Set rs = conn.Execute(sql)

    ' gather final grades into array
    Dim finalGrades() As Double
    Dim count As Long, total As Double, i As Long
    count = 0: total = 0

    Do While Not rs.EOF
        ReDim Preserve finalGrades(count)
        finalGrades(count) = rs("FinalGrade")
        total = total + finalGrades(count)
        count = count + 1
        rs.MoveNext
    Loop

    rs.Close: conn.Close

    If count = 0 Then
        MsgBox "No grades found!", vbInformation
        Exit Sub
    End If

    ' calculate stats
    Dim avg As Double, minG As Double, maxG As Double, stdev As Double
    Dim sumSqDiff As Double: sumSqDiff = 0
    minG = finalGrades(0): maxG = finalGrades(0)

    For i = 0 To count - 1
        If finalGrades(i) < minG Then minG = finalGrades(i)
        If finalGrades(i) > maxG Then maxG = finalGrades(i)
    Next i

    avg = total / count

    For i = 0 To count - 1
        sumSqDiff = sumSqDiff + (finalGrades(i) - avg) ^ 2
    Next i
    stdev = Sqr(sumSqDiff / count)

    ' create histogram data
    Dim bins() As String: bins = Split("0-49,50-59,60-69,70-79,80-89,90-100", ",")
    Dim binCounts() As Long: ReDim binCounts(UBound(bins))

    For i = 0 To count - 1
        Select Case finalGrades(i)
            Case Is < 50: binCounts(0) = binCounts(0) + 1
            Case 50 To 59: binCounts(1) = binCounts(1) + 1
            Case 60 To 69: binCounts(2) = binCounts(2) + 1
            Case 70 To 79: binCounts(3) = binCounts(3) + 1
            Case 80 To 89: binCounts(4) = binCounts(4) + 1
            Case Is >= 90: binCounts(5) = binCounts(5) + 1
        End Select
    Next i

    ' output to worksheet
    Dim ws As Worksheet, chartObj As ChartObject
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "ReportChart"
    ws.Cells(1, 1).Value = "Grade Range"
    ws.Cells(1, 2).Value = "Student Count"

    For i = 0 To UBound(bins)
        ws.Cells(i + 2, 1).Value = bins(i)
        ws.Cells(i + 2, 2).Value = binCounts(i)
    Next i

    ' insert chart
    Set chartObj = ws.ChartObjects.Add(Left:=150, Width:=400, Top:=10, Height:=300)
    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("A1:B" & UBound(bins) + 2)
        .HasTitle = True
        .ChartTitle.Text = "Histogram of Final Grades"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Grade Range"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Number of Students"
    End With

    ' create word report
    Dim wordApp As Object, doc As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add

    ' summary description of project
    doc.Content.InsertAfter "Final Grade Report" & vbCrLf & vbCrLf
    doc.Content.InsertAfter "This report was generated by the Grading Application created in Excel VBA. " & _
                            "It calculates each student’s final grade based on assignment and exam weights, " & _
                            "groups the results into grade ranges, and produces a histogram summarizing the class distribution." & vbCrLf & vbCrLf

    ' insert stats
    doc.Content.InsertAfter "Final Grade Summary:" & vbCrLf
    doc.Content.InsertAfter "Average: " & Format(avg, "0.00") & "%" & vbCrLf
    doc.Content.InsertAfter "Minimum: " & Format(minG, "0.00") & "%" & vbCrLf
    doc.Content.InsertAfter "Maximum: " & Format(maxG, "0.00") & "%" & vbCrLf
    doc.Content.InsertAfter "Standard Deviation: " & Format(stdev, "0.00") & "%" & vbCrLf & vbCrLf

    ' paste the chart
    chartObj.Chart.ChartArea.Copy
    doc.Content.InsertParagraphAfter
    wordApp.Selection.Paste

    ' notify user
    MsgBox "Word report generated successfully!", vbInformation
End Sub
'if user selects import data button
Private Sub cmdImportData_Click()
 'hide current form and display the appropriate form
    Me.Hide
    frmImport.Show
End Sub

Private Sub cmdViewGrades_Click()
'check database has been loaded first
   If txtFilePath.Text = "" Then
        MsgBox "Please import the database first!", vbExclamation
        Exit Sub
    End If
    'hide current form and display the appropriate form
    Me.Hide
    frmViewGrades.Show
End Sub

Private Sub UserForm_Initialize()
'instructions for user on how to use application
    MsgBox "Welcome to the Grading Application!" & vbCrLf & vbCrLf & _
           "Click 'Import Data' to load the registrar.mdb file." & vbCrLf & _
           "Then explore: Display Averages, Generate Charts, View Grades, or Create a Word Report.", _
           vbInformation, "Instructions"
End Sub
