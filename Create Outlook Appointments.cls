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
' Updated Feb 2016 to pull data from Maximo instead of assyst, plus a variety of enhancements.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Const C_WONUM = 1
Const C_DESCRIPTION = 2
Const C_OWNER = 3
Const C_OWNERGROUP = 4
Const C_STATUS = 5
Const C_ChgType = 6
Const C_ChgCat = 7
Const C_SCHEDSTART = 8
Const C_SCHEDFINISH = 9
Const C_CINAME = 10
Const C_LONG_DESC = 11
Const C_CHGREASON = 12
Const C_NOT_IMPL_EFFECT = 13
Const C_LOGS = 15
Const LAST_UPDATE_DATE_CELL = "E1"



Private Sub cmdCreateApps_Click()

    Dim oApp As Object
    Dim oNameSpace As Namespace
    Dim oFolder As Outlook.MAPIFolder
    Dim MyItem As Outlook.AppointmentItem
    Dim ChgType As String
    Dim ChgCat As String
    Dim ApptNum As Long
    Dim STATUS As String
    Dim NewItem As Boolean


    '''''''''''''''''''''''''''''''''''
    ' Testing starting date for query
    '''''''''''''''''''''''''''''''''''

    Set oApp = New Outlook.Application
    Set oNameSpace = oApp.GetNamespace("MAPI")
    'Set oFolder = oNameSpace.Folders("Mailbox - changecontrol@enmax.com").Folders("Calendar")
    'Set oFolder = oNameSpace.Folders("Public Folders").Folders("All Public Folders").Folders("Corporate").Folders("Information Technology").Folders("Forward Schedule of Change")
    Set oFolder = oNameSpace.PickFolder()
    

    Dim EndRow As Integer
    Dim CurrentRow As Integer


    ActiveWorkbook.Connections("MITPRD - FSC Change Details").ODBCConnection.BackgroundQuery = False
    ActiveWorkbook.Connections("MITPRD - FSC Change Details").Refresh

    usrProgress.txtAdded = 0
    usrProgress.txtUpdated = 0

    EndRow = LastRow("Report")
    
    usrProgress.Show vbModeless
    
    For CurrentRow = 3 To EndRow
    
    usrProgress.txtTotal = EndRow - 2
    usrProgress.txtCurrent = CurrentRow - 2
    usrProgress.txtPercent = Round((CurrentRow - 2) / (EndRow - 2), 2) * 100
    
    usrProgress.Repaint
    
    On Error Resume Next
    

    ' Check to see if the item already exists in the calendar
        Set MyItem = oFolder.Items.Find("[Subject] = """ & Sheets("Report").Cells(CurrentRow, C_WONUM) & """")
        
        If MyItem Is Nothing Then
            Set MyItem = oFolder.Items.Add
            usrProgress.txtAdded = usrProgress.txtAdded + 1
            NewItem = True
            Err.Clear
        Else
            usrProgress.txtUpdated = usrProgress.txtUpdated + 1
            NewItem = False
        End If
        
        ChgType = Sheets("Report").Cells(CurrentRow, C_ChgType)
        ChgCat = Sheets("Report").Cells(CurrentRow, C_ChgCat)
        STATUS = Sheets("Report").Cells(CurrentRow, C_STATUS)
        
        If ChgCat <> "" Then
            Select Case ChgCat
                Case "Minor"
                    MyItem.Categories = "Green Category"
                Case "Pivot DC"
                    MyItem.Categories = "Green Category"
                Case "Shaw Business"
                    MyItem.Categories = "Green Category"
                Case "Standard"
                    MyItem.Categories = "Green Category"
                Case "Significant"
                    MyItem.Categories = "Yellow Category"
                Case "Business Requirement"
                    MyItem.Categories = "Red Category"
                Case "Incident"
                    MyItem.Categories = "Red Category"
                Case Else
                    MyItem.Categories = ""
            End Select
        End If
        
        Select Case STATUS
            Case "CLOSE"
                MyItem.Categories = "Closed Change"
            Case "CAN"
                MyItem.Categories = "Closed Change"
            Case "COMP"
                MyItem.Categories = "Closed Change"
            Case "REVIEW"
                MyItem.Categories = "Closed Change"
        End Select
        
        'MyItem.Location = ChgType & " - " & ChgCat & " - (" & Sheets("Report").Cells(CurrentRow, C_STATUS) & ")"
        MyItem.Location = Sheets("Report").Cells(CurrentRow, C_DESCRIPTION) & " - " & ChgType & ", " & ChgCat & " (" & Sheets("Report").Cells(CurrentRow, C_STATUS) & ")"
        
        MyItem.Subject = Sheets("Report").Cells(CurrentRow, C_WONUM)
        MyItem.Start = Sheets("Report").Cells(CurrentRow, C_SCHEDSTART)
        MyItem.End = Sheets("Report").Cells(CurrentRow, C_SCHEDFINISH)
        MyItem.ReminderSet = False
        MyItem.Body = "Change #:" & vbTab & vbTab & Sheets("Report").Cells(CurrentRow, C_WONUM) & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Change Owner:" & vbTab & vbTab & Sheets("Report").Cells(CurrentRow, C_OWNER)
        MyItem.Body = MyItem.Body & vbCrLf & "Change Owner Group:" & vbTab & vbTab & Sheets("Report").Cells(CurrentRow, C_OWNERGROUP)
        MyItem.Body = MyItem.Body & vbCrLf & "Scheduled Start:" & vbTab & Sheets("Report").Cells(CurrentRow, C_SCHEDSTART)
        MyItem.Body = MyItem.Body & vbCrLf & "Scheduled End:" & vbTab & vbTab & Sheets("Report").Cells(CurrentRow, C_SCHEDFINISH)
        MyItem.Body = MyItem.Body & vbCrLf & "Affected CI:" & vbTab & vbTab & Sheets("Report").Cells(CurrentRow, C_CINAME)
        MyItem.Body = MyItem.Body & vbCrLf & "Current status:" & vbTab & vbTab & Sheets("Report").Cells(CurrentRow, C_STATUS)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Summary:" & vbCrLf & Sheets("Report").Cells(CurrentRow, C_DESCRIPTION)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Reason for Change:" & vbCrLf & Sheets("Report").Cells(CurrentRow, C_CHGREASON)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Effect of not Implementing:" & vbCrLf & Sheets("Report").Cells(CurrentRow, C_NOT_IMPL_EFFECT)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Details:" & vbCrLf & Sheets("Report").Cells(CurrentRow, C_LONG_DESC)
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "-----------------------------------------------"
        MyItem.Body = MyItem.Body & vbCrLf
        MyItem.Body = MyItem.Body & vbCrLf & "Work Logs:" & vbCrLf & Sheets("Report").Cells(CurrentRow, C_LOGS)
        MyItem.Body = MyItem.Body & vbCrLf
        
        MyItem.Save
    
    Next CurrentRow

    Set oApp = Nothing
    Set oNameSpace = Nothing
    Set oFolder = Nothing
    Set MyItem = Nothing
    
    MsgBox "Update complete", vbOKOnly, "Complete"
    
    Unload usrProgress
    
    Sheets("Report").Range(LAST_UPDATE_DATE_CELL) = Format(Now(), "dd-mm-yyyy hh:mm:ss")

End Sub
