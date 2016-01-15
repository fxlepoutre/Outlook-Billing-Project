Attribute VB_Name = "TimeCounter"

Function FindDuration(sInputCategory As String, datInputStartDate As Date, datInputEndDate As Date)

    Dim datStart, datEnd As Date
    Dim objCalendar As Outlook.Folder
    Dim objItems As Outlook.Items
    Dim objAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    Dim dblDuration As Double

    Set objCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set objItems = objCalendar.Items
    
    ' Include recurring appointments.
    objItems.IncludeRecurrences = True

    ' Format dates as Outlook espects them.
    datStart = Format(datInputStartDate, "ddddd h:nn AM/PM")
    datEnd = Format(datInputEndDate, "ddddd h:nn AM/PM")
    
    ' Construct a filter.
    strRestriction = "[Categories] = '" & sInputCategory & "' AND [Show Time As] = 'Busy' AND [Start] >= '" & datStart & "' AND [End] <= '" & datEnd & "'"
    Debug.Print strRestriction
    ' Restrict the Items collection.
    Set objItems = objItems.Restrict(strRestriction)
        
    ' Sort and print the final results.
    objItems.Sort "[Start]"
    dblDuration = 0
    For Each objAppt In objItems
        dblDuration = dblDuration + objAppt.Duration
        Debug.Print objAppt.Start, objAppt.Duration, objAppt.End, objAppt.Categories, objAppt.Subject
    Next
    Debug.Print "Total duration: ", Format(-(-Int(dblDuration / 60)), "#0") & " hours "; Format(dblDuration / 60 / 24, "nn") & " minutes"
    FindDuration = dblDuration

End Function


Sub CountTime()
    Dim objNameSpace As NameSpace
    Dim objCategory As Category
    Dim strOutput As String
    Dim dblDuration As Double
    Dim dblTotalDuration As Double
    Dim objView As Outlook.View
    Dim datStartDate As Date
    Dim datEndDate As Date
    Dim strSearchString As String
    
    ' Check if the current view is a table view.
    If Application.ActiveExplorer.CurrentView.ViewType = olTableView Then
        ' Set the view to the active view.
        Set objView = Application.ActiveExplorer.CurrentView
        Debug.Print "Current view : ", objView.Name, objView.Filter
        ' Search the start date of the filter.
        strSearchString = """urn:schemas:calendar:dtstart"" >= '"
        If InStr(objView.Filter, strSearchString) > 0 Then
            datStartDate = CDate(Mid(objView.Filter, InStr(objView.Filter, strSearchString) + Len(strSearchString), 10))
        Else
            MsgBox "The current view should be filtered with a start date, before which it does not include any items (start date on or after dd/mm/yyyy). Please use ""View Settings"" in ribbon, ""Filter..."" section, ""Advanced"" tab."
            Exit Sub
        End If
        Debug.Print "Start date:", datStartDate
        ' Search the end date of the filter.
        strSearchString = """urn:schemas:calendar:dtend"" <= '"
        If InStr(objView.Filter, strSearchString) > 0 Then
            datEndDate = CDate(Mid(objView.Filter, InStr(objView.Filter, strSearchString) + Len(strSearchString), 10))
        Else
            MsgBox "The current view should be filtered with an end date, after which it does not include any items (end date on or before dd/mm/yyyy).  Please use ""View Settings"" in ribbon, ""Filter..."" section, ""Advanced"" tab."
            Exit Sub
        End If
        Debug.Print "End date:", datEndDate
    Else
        MsgBox "The current view should be a filtered table view. Please use view settings in ribbon."
        Exit Sub
    End If

    
'    ' Asks for limit dates
'    datStartDate = CDate(InputBox(Prompt:="Please input the date BEFORE which the filter does not include items. Use the following format : dd/mm/yyyy. Items on this date will be included.", Title:="Filter start date", Default:="25/04/2011"))
'    datEndDate = CDate(InputBox(Prompt:="Please input the date AFTER which the filter does not include items. Use the following format : dd/mm/yyyy. Items on this date will be included.", Title:="Filter end date", Default:="27/05/2011"))
    
    ' Obtain a NameSpace object reference.
    Set objNameSpace = Application.GetNamespace("MAPI")
    
    ' Initialize total duration.
    dblTotalDuration = 0
    
    ' Check if the Categories collection for the Namespace contains one or more Category objects.
    If objNameSpace.Categories.Count > 0 Then
        ' Enumerate the Categories collection.
        For Each objCategory In objNameSpace.Categories
            ' Remove the personnal category
            If objCategory.Name <> "0- Personnal" Then
                dblDuration = FindDuration(objCategory.Name, datStartDate, datEndDate)
                If dblDuration > 0 Then
                    dblTotalDuration = dblTotalDuration + dblDuration
                    strOutput = strOutput & objCategory.Name & ": " & Format(-(-Int(dblDuration / 60)), "#0") & " hours " & Format(dblDuration / 60 / 24, "nn") & " minutes" & vbCrLf
                End If
            End If
        Next
        strOutput = strOutput & "Total: " & Format(-(-Int(dblTotalDuration / 60)), "#0") & " hours " & Format(dblTotalDuration / 60 / 24, "nn") & " minutes (" & dblTotalDuration / 60 / 8 & " days)" & vbCrLf
    End If
    
    ' Display the output string.
    MsgBox strOutput
    
    ' Clean up.
    Set objView = Nothing
    Set objCategory = Nothing
    Set objNameSpace = Nothing
    
End Sub
