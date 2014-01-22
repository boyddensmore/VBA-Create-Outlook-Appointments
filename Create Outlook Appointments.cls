VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Purpose:  Load data pulled from assyst into an Outlook calendar for easy viewing of
'           past/upcoming changes.
' Author:   Boyd Densmore, 2011
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub cmdCreateApps_Click()

    Dim oApp As Object
    Dim oNameSpace As Namespace
    Dim oFolder As Outlook.MAPIFolder
    Dim MyItem As Outlook.AppointmentItem
    Dim Category As String
    Dim ApptNum As Long
    Dim Status As String
    Dim NewItem As Boolean

    Set oApp = New Outlook.Application
    Set oNameSpace = oApp.GetNamespace("MAPI")
    'Set oFolder = oNameSpace.Folders("Mailbox - changecontrol@enmax.com").Folders("Calendar")
    Set oFolder = oNameSpace.Folders("Public Folders - bdensmore@enmax.com").Folders("All Public Folders").Folders("Corporate").Folders("Information Technology").Folders("Forward Schedule of Change")

    Dim EndRow As Integer
    Dim CurrentRow As Integer

    usrProgress.txtAdded = 0
    usrProgress.txtUpdated = 0

    EndRow = LastRow("Data Prep")
    
    usrProgress.Show vbModeless
    
    For CurrentRow = 2 To EndRow
    
    usrProgress.txtTotal = EndRow - 1
    usrProgress.txtCurrent = CurrentRow - 1
    usrProgress.txtPercent = Round((CurrentRow - 1) / (EndRow - 1), 2) * 100
    
    usrProgress.Repaint
    
    On Error Resume Next
    
        If oFolder.Items("R" & Sheets("Data Prep").Cells(CurrentRow, 5)) = "R" & Sheets("Data Prep").Cells(CurrentRow, 5) Then
            If Err.Number <> 0 Then
                Set MyItem = oFolder.Items.Add
                usrProgress.txtAdded = usrProgress.txtAdded + 1
                NewItem = True
                Err.Clear
            Else
                Set MyItem = oFolder.Items("R" & Sheets("Data Prep").Cells(CurrentRow, 5))
                usrProgress.txtUpdated = usrProgress.txtUpdated + 1
                NewItem = False
            End If
        End If
    
        If Err.Number = 91 Then
            MsgBox "You don't have permissions to add/update FSC items.  Please contact the Service Desk."
        End If
    
        Category = Sheets("Data Prep").Cells(CurrentRow, 14)
        Status = Sheets("Data Prep").Cells(CurrentRow, 15)
        
        If Category <> "" Then
            Select Case Category
                Case "RFC MINOR"
                    MyItem.Categories = "Green Category"
                Case "RFC STANDARD"
                    MyItem.Categories = "Green Category"
                Case "RFC INFORMATION"
                    MyItem.Categories = "Green Category"
                Case "RFC MAJOR"
                    MyItem.Categories = "Red Category"
                Case "RFC MEDIUM"
                    MyItem.Categories = "Orange Category"
                Case "RFC URGENT"
                    MyItem.Categories = "Red Category"
                Case "RFC EMERGENCY"
                    MyItem.Categories = "Red Category"
                Case Else
                    MyItem.Categories = ""
                    Category = "UNCATEGORIZED"
            End Select
        End If
        
        Select Case Status
            Case "o"
                MyItem.Location = Category & " - " & Sheets("Data Prep").Cells(CurrentRow, 6) & " (Open)"
            Case Else
                MyItem.Location = Category & " - " & Sheets("Data Prep").Cells(CurrentRow, 6) & " (Closed)"
                MyItem.Categories = "Closed Change"
        End Select
        
        MyItem.Subject = "R" & Sheets("Data Prep").Cells(CurrentRow, 5)
        MyItem.Start = Sheets("Data Prep").Cells(CurrentRow, 3)
        MyItem.End = Sheets("Data Prep").Cells(CurrentRow, 4)
        MyItem.ReminderSet = False
        MyItem.Body = "Ticket number:" & vbTab & vbTab & "R" & Sheets("Data Prep").Cells(CurrentRow, 5) & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Logged By:" & vbTab & vbTab & Sheets("Data Prep").Cells(CurrentRow, 7)
        MyItem.Body = MyItem.Body & vbCrLf & "Group:" & vbTab & vbTab & vbTab & Sheets("Data Prep").Cells(CurrentRow, 16)
        MyItem.Body = MyItem.Body & vbCrLf & "Date Logged:" & vbTab & vbTab & Sheets("Data Prep").Cells(CurrentRow, 1)
        MyItem.Body = MyItem.Body & vbCrLf & "Scheduled Start:" & vbTab & Sheets("Data Prep").Cells(CurrentRow, 3)
        MyItem.Body = MyItem.Body & vbCrLf & "Scheduled End:" & vbTab & vbTab & Sheets("Data Prep").Cells(CurrentRow, 4)
        MyItem.Body = MyItem.Body & vbCrLf & "Current stage:" & vbTab & vbTab & Sheets("Data Prep").Cells(CurrentRow, 6)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Description:" & vbCrLf & Sheets("Data Prep").Cells(CurrentRow, 13)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Justification:" & vbCrLf & Sheets("Data Prep").Cells(CurrentRow, 8)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Risk:" & vbCrLf & Sheets("Data Prep").Cells(CurrentRow, 9)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Implementation:" & vbCrLf & Sheets("Data Prep").Cells(CurrentRow, 10)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Backout:" & vbCrLf & Sheets("Data Prep").Cells(CurrentRow, 11)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Testing:" & vbCrLf & Sheets("Data Prep").Cells(CurrentRow, 12)
        
        MyItem.Save
    
    Next CurrentRow

    Set oApp = Nothing
    Set oNameSpace = Nothing
    Set oFolder = Nothing
    Set MyItem = Nothing
    
    MsgBox "Update complete", vbOKOnly, "Complete"
    
    Unload usrProgress

End Sub

Private Sub cmdPrepData_Click()

    Dim EndRow As Integer
    Dim CurrentRow As Integer
    Dim CurrentCol As Integer
    Dim CurrRow As Integer
    Dim FirstUsedRow As Integer
    Dim LastUsedRow As Integer
    Dim rFound As Range
    Dim sRef As String
   
    FirstUsedRow = 2
    
    Sheets("Data Prep").Range("A2:" & SheetBoundaries("Data Prep")(2) & SheetBoundaries("Data Prep")(1)).Clear
    
    On Error Resume Next
    
    EndRow = LastRow("Report")
    
    For CurrentRow = 3 To EndRow
    
        For CurrentCol = 1 To 13
            Sheets("Data Prep").Cells(CurrentRow - 1, CurrentCol) = Replace(Sheets("Report").Cells(CurrentRow, CurrentCol), Chr(13), "")
        Next CurrentCol
        
        Sheets("Data Prep").Range("M" & CurrentRow - 1) = Replace(Sheets("Report").Range("N" & CurrentRow) & Sheets("Report").Range("O" & CurrentRow) & Sheets("Report").Range("P" & CurrentRow) & Sheets("Report").Range("Q" & CurrentRow) & Sheets("Report").Range("R" & CurrentRow) & Sheets("Report").Range("S" & CurrentRow) & Sheets("Report").Range("T" & CurrentRow) & Sheets("Report").Range("U" & CurrentRow), Chr(13), "")

        Sheets("Data Prep").Range("N" & CurrentRow - 1) = Sheets("Report").Range("W" & CurrentRow)

        Sheets("Data Prep").Range("O" & CurrentRow - 1) = Sheets("Report").Range("V" & CurrentRow)
        
        Sheets("Data Prep").Range("P" & CurrentRow - 1) = Sheets("Report").Range("Y" & CurrentRow)
    
    Next CurrentRow


End Sub
