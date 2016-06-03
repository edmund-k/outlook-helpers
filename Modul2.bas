Attribute VB_Name = "Modul2"
' Extracts the attendee list from an appointment,
' puts the attendees into categories by response status (accepted, tentative, declined, no response) and
' opens up a formatted email printing out the response status of the attendees.
Sub GetAttendeeList()
      
    Dim objApp As Outlook.Application
    Dim objItem As Object
    Dim objSelection As Selection
    Dim objAttendees As Outlook.Recipients
    Dim objAttendeeReq As String
    Dim objAttendeeOpt As String
    Dim objAttendeeAcc As String
    Dim objAttendeeTen As String
    Dim objAttendeeDec As String
    Dim objAttendeeNor As String
    Dim objOrganizer As String
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim strSubject As String
    Dim strLocation As String
    Dim strNotes As String
    Dim strMeetStatus As String
    Dim strCopyData As String
      
    On Error Resume Next
      
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    Set objSelection = objApp.ActiveExplorer.Selection
    Set objAttendees = objItem.Recipients
      
    On Error GoTo EndClean:
      
    ' Check edge cases with none or more than one item open. Only one opened item will be accepted.
    Select Case objSelection.Count
        Case 0
            MsgBox "No meeting was opened.  Please open the meeting to print."
            GoTo EndClean:
        Case Is > 1
            MsgBox "Too many items.  Just select one!"
            GoTo EndClean:
    End Select
      
    ' Is it an appointment?
    If objItem.Class <> 26 Then
        MsgBox "You need to open the meeting."
        GoTo EndClean:
    End If
      
    ' Get the data!
    dtStart = objItem.Start
    dtEnd = objItem.End
    strSubject = objItem.Subject
    strLocation = objItem.Location
    strNotes = objItem.Body
    objOrganizer = objItem.Organizer
    objAttendeeReq = ""
    objAttendeeOpt = ""
    objAttendeeAcc = ""
    objAttendeeTen = ""
    objAttendeeDec = ""
    objAttendeeNor = ""
      
    ' Get the attendee list and put the attendees into categories.
    For x = 1 To objAttendees.Count
        strMeetStatus = ""
        Select Case objAttendees(x).MeetingResponseStatus
            Case 0
                strMeetStatus = "No Response"
                objAttendeeNor = objAttendeeNor & objAttendees(x).Name & "; "
            Case 1
                strMeetStatus = "Organizer"
            Case 2
                strMeetStatus = "Tentative"
                objAttendeeTen = objAttendeeTen & objAttendees(x).Name & "; "
            Case 3
                strMeetStatus = "Accepted"
                objAttendeeAcc = objAttendeeAcc & objAttendees(x).Name & "; "
            Case 4
                strMeetStatus = "Declined"
                objAttendeDec = objAttendeeDec & objAttendees(x).Name & "; "
        End Select
       
        If objAttendees(x).Type = olRequired Then
            objAttendeeReq = objAttendeeReq & objAttendees(x).Name & "; required" & "; " & strMeetStatus & vbCrLf
        Else
            objAttendeeOpt = objAttendeeOpt & objAttendees(x).Name & "; optional" & "; " & strMeetStatus & vbCrLf
        End If
    Next
       
    ' Open up a formated email printing out the response status of the attendees.
    strCopyData = _
        "Subject: " & strSubject & vbCrLf & _
        "Location: " & strLocation & vbCrLf & _
        "Start: " & dtStart & vbCrLf & _
        "End: " & dtEnd & vbCrLf & vbCrLf & _
        "Required: " & vbCrLf & objAttendeeReq & vbCrLf & vbCrLf & _
        "Optional: " & vbCrLf & objAttendeeOpt & vbCrLf & vbCrLf & _
        "Accepted: " & objAttendeeAcc & vbCrLf & vbCrLf & _
        "Tentative: " & objAttendeeTen & vbCrLf & vbCrLf & _
        "Declined: " & objAttendeeDec & vbCrLf & vbCrLf & _
        "No Response: " & objAttendeeNor & vbCrLf & vbCrLf
    Set ListAttendees = Application.CreateItem(olMailItem)
        ListAttendees.Body = strCopyData
        ListAttendees.Display
        
' Clean up variables and free up memory.
EndClean:
    Set objApp = Nothing
    Set objItem = Nothing
    Set objSelection = Nothing
    Set objAttendees = Nothing

End Sub

