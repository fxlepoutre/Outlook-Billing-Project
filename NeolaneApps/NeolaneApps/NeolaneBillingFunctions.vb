Module nlBillingFunctions

    ''' <summary>
    ''' Returns the invoicing status of a task.
    ''' </summary>
    ''' <param name="objInputAppointment">Outlook task.</param>
    ''' <returns>A string containing the invoice status.</returns>
    ''' <remarks></remarks>
    Private Function GetInvoicingStatus(ByVal objInputAppointment As Outlook.AppointmentItem) As String
        'Task is to invoice only if its start date is not in the future, if its status is "Busy" or "Out Of Office" and if it is categorized.
        If objInputAppointment.Start <= Now() And (objInputAppointment.BusyStatus = Outlook.OlBusyStatus.olBusy Or objInputAppointment.BusyStatus = Outlook.OlBusyStatus.olOutOfOffice) Then
            GetInvoicingStatus = "To invoice"
        ElseIf objInputAppointment.Start > Now() Or (objInputAppointment.BusyStatus = Outlook.OlBusyStatus.olTentative Or objInputAppointment.BusyStatus = Outlook.OlBusyStatus.olFree) Or IsNothing(objInputAppointment.Categories) Then
            GetInvoicingStatus = "Do not invoice"
        Else
            GetInvoicingStatus = "Cannot get invoicing status"
        End If
    End Function

    ''' <summary>
    ''' Returns the busy status text of a task.
    ''' </summary>
    ''' <param name="intInputBusyStatus">Outlook-based value for task status.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetBusyStatus(ByVal intInputBusyStatus As Integer) As String
        'Default Outlook meanings.
        Select Case intInputBusyStatus
            Case 0
                GetBusyStatus = "Free"
            Case 1
                GetBusyStatus = "Tentative"
            Case 2
                GetBusyStatus = "Busy"
            Case 3
                GetBusyStatus = "Out of office"
            Case Else
                GetBusyStatus = "Status unknown"
        End Select
    End Function

    ''' <summary>
    ''' Procedure to list all Outlook tasks between two dates.
    ''' </summary>
    ''' <param name="inputDatStart">Date on or after listing should be done.</param>
    ''' <param name="inputDatEnd">Date on or before listing should be done.</param>
    ''' <remarks></remarks>
    Public Sub ListTasks(ByVal inputDatStart As Date, ByVal inputDatEnd As Date)

        Dim datStart, datEnd As Date
        Dim objOutlookApp As Outlook.Application
        Dim objNameSpace As Outlook.NameSpace
        Dim objCalendar As Outlook.Folder
        Dim objItems As Outlook.Items
        Dim objAppt As Outlook.AppointmentItem
        Dim strRestriction As String
        Dim intLineCount As Integer
        Dim intColumnCount As Integer
        Dim objExcelApp As Excel.Application
        Dim objBook As Excel.Workbook
        Dim objSheet1 As Excel.Worksheet
        Dim objSheet2 As Excel.Worksheet

        'Start a new workbook in Excel
        objExcelApp = CreateObject("Excel.Application")
        objBook = objExcelApp.Workbooks.Add
        objSheet1 = objBook.Worksheets(1)
        objSheet2 = objBook.Worksheets(2)
        objExcelApp.DisplayAlerts = False
        objBook.Worksheets(3).Delete()
        objExcelApp.DisplayAlerts = True
        objSheet1.Name = "Data"
        objSheet2.Name = "PivotTable"

        'Show the Excel sheet.
        objExcelApp.Visible = True

        objOutlookApp = New Outlook.Application()
        objNameSpace = objOutlookApp.GetNamespace("MAPI")

        objCalendar = objNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
        objItems = objCalendar.Items

        'Include recurring appointments.
        objItems.IncludeRecurrences = True
        'Sort the Items collection
        objItems.Sort("[Start]")

        'Format dates as Outlook espects them.
        datStart = inputDatStart
        datEnd = DateAdd(DateInterval.Day, 1, inputDatEnd)

        'Construct a filter.
        strRestriction = "[Start] >= '" & datStart & "' AND [End] <= '" & datEnd & "'"

        'Restrict the Items collection according to the previous filter.
        objItems = objItems.Restrict(strRestriction)

        intColumnCount = 1
        objSheet1.Range("A1:H1").Value = {"StartDate", "EndDate", "Label", "Duration (hours)", "Duration (days)", "Categories", "Busy Status", "Invoicing"}

        'Initialize task count.
        intLineCount = 1

        'Place items in our object.
        For Each objAppt In objItems
            'Add item in Excel
            intLineCount = intLineCount + 1
            objSheet1.Range("A" & intLineCount & ":H" & intLineCount).Value = {
                Format(objAppt.Start, "MM/dd/yyyy HH:mm:ss"),
                Format(objAppt.End, "MM/dd/yyyy HH:mm:ss"),
                Format(objAppt.Start, "yyyy/MM/dd") & " - " & objAppt.Subject,
                objAppt.Duration / 60,
                objAppt.Duration / 60 / 8,
                objAppt.Categories,
                GetBusyStatus(objAppt.BusyStatus),
                GetInvoicingStatus(objAppt)
            }
        Next

        objSheet1.Range("A1:H" & intLineCount).EntireColumn.AutoFit()

        objSheet1.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, objSheet1.Range("A1:H" & intLineCount), , Excel.XlYesNoGuess.xlYes).Name = "ListTasks"
        objSheet1.ListObjects("ListTasks").TableStyle = "TableStyleLight2"

        Dim objTable As Excel.PivotTable
        ' Create the PivotTable object.
        objBook.PivotCaches.Create( _
                SourceType:=Excel.XlPivotTableSourceType.xlDatabase, _
                SourceData:="ListTasks", _
                Version:=Excel.XlPivotTableVersionList.xlPivotTableVersion14 _
            ).CreatePivotTable( _
                TableDestination:=objSheet2.Name & "!R1C1", _
                TableName:="PivotTableBilling", _
                DefaultVersion:=Excel.XlPivotTableVersionList.xlPivotTableVersion14)
        objTable = objSheet2.PivotTables("PivotTableBilling")

        'Add dimensions, sorts and filters to pivota table.
        With objTable.PivotFields("Categories")
            .Orientation = Excel.XlPivotFieldOrientation.xlRowField
            .Position = 1
            Try
                .PivotItems("Holiday").Visible = False
            Catch ex As Exception
            End Try
            Try
                .PivotItems("(blank)").Visible = False
            Catch ex As Exception
            End Try
        End With
        With objTable.PivotFields("Label")
            .Orientation = Excel.XlPivotFieldOrientation.xlRowField
            .Position = 2
        End With
        With objTable.PivotFields("Categories")
            .ShowDetail = False
        End With
        With objTable.PivotFields("Busy Status")
            .Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .Position = 1
        End With
        With objTable.PivotFields("Invoicing")
            .Orientation = Excel.XlPivotFieldOrientation.xlPageField
            .Position = 2
        End With
        With objTable.PivotFields("StartDate")
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            .Position = 1
        End With

        'Group table by week, every week starting on Saturday (day number 6), because reporting and billing is made on Friday evening.
        Dim datGroupStartInterval As Integer
        Dim datGroupStartWeekDay As DayOfWeek = DayOfWeek.Saturday
        If datStart.DayOfWeek >= datGroupStartWeekDay Then
            datGroupStartInterval = datGroupStartWeekDay - datStart.DayOfWeek
        ElseIf datStart.DayOfWeek < datGroupStartWeekDay Then
            datGroupStartInterval = datGroupStartWeekDay - datStart.DayOfWeek - 7
        End If
        'Try to group dates by weeks. Known causes for exception thrown: List of tasks is empty, dates are not stored as dates in objSheet1.
        Try
            objSheet2.Range("B5").Group(By:=7, Periods:={False, False, False, True, False, False, False}, Start:=DateAdd(DateInterval.Day, datGroupStartInterval, datStart))
        Catch ex As Exception
        End Try

        'Add data
        objTable.AddDataField(objTable.PivotFields("Duration (hours)"), "Total duration (hours)", Excel.XlConsolidationFunction.xlSum)

        'Add info on top of sheet.
        objSheet2.Range("C1").Value = "Report start date (included): " & Format(inputDatStart, "dd MMM yyyy")
        objSheet2.Range("C2").Value = "Report end date (included): " & Format(inputDatEnd, "dd MMM yyyy")
        objSheet2.Range("C1:E2").Merge(True)
        objSheet2.Range("C1:E2").HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

        'Show result sheet
        objSheet2.Select()
    End Sub
End Module