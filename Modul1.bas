Attribute VB_Name = "Modul1"
' Sets the according category for a selected mail item, clearing all other categories.

Sub StatusNextAction()
    
    Dim mi As MailItem
    Dim item As Object
    
    For Each item In Application.ActiveExplorer.Selection
        Set mi = item
        mi.Categories = "S/1Next Action"
        If Not mi.IsMarkedAsTask Then
            mi.ClearTaskFlag
        End If
        mi.Save
    Next
    
End Sub


Sub StatusAction()
    
    Dim mi As MailItem
    Dim item As Object
    
    For Each item In Application.ActiveExplorer.Selection
        Set mi = item
        mi.Categories = "S/2Action"
        If Not mi.IsMarkedAsTask Then
            mi.ClearTaskFlag
        End If
        mi.Save
    Next
    
End Sub


Sub StatusSomeday()

    Dim mi As MailItem
    Dim item As Object
    
    For Each item In Application.ActiveExplorer.Selection
        Set mi = item
        mi.Categories = "S/3Someday"
        mi.ClearTaskFlag
        mi.Save
    Next
    
End Sub


Sub StatusWaitingOn()
    
    Dim mi As MailItem
    Dim item As Object
    
    For Each item In Application.ActiveExplorer.Selection
        Set mi = item
        mi.Categories = "S/4Waiting On"
        If Not mi.IsMarkedAsTask Then
            mi.ClearTaskFlag
        End If
        mi.Save
    Next
    
End Sub


Sub StatusFinished()
    
    Dim mi As MailItem
    Dim item As Object
    
    For Each item In Application.ActiveExplorer.Selection
        Set mi = item
        mi.Categories = "S/5Finished"
        mi.TaskCompletedDate = Date
        mi.Save
    Next
    
End Sub


