Public Class NeolaneRibbon

    Structure nlBillingPeriod
        Public StartDate As Date
        Public BillingDate As Date
        Public PeriodName As String
        Public PeriodOrder As Integer
    End Structure

    Private Const XML_DOCUMENT_LOCATION As String = "C:\NeolaneBillingDates.xml"

    ''' <summary>
    ''' Read XML document and parse billing dates to create billing periods.
    ''' </summary>
    ''' <returns>An array of the billing periods.</returns>
    ''' <remarks></remarks>
    Private Function GetBillingPeriodsFromXML() As nlBillingPeriod()

        Dim BillingPeriod() As nlBillingPeriod
        Dim xmlDocument As Xml.XmlDocument
        Dim xmlNodeList As Xml.XmlNodeList
        Dim xmlNode As Xml.XmlNode
        Dim datPeriodStartDate As Date
        Dim intNumberOfXmlNodes As Integer = 0

        'Load XML document from xmlDocumentLocation, defined as a const previously.
        xmlDocument = New Xml.XmlDocument()
        Try
            xmlDocument.Load(XML_DOCUMENT_LOCATION)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try

        'Select nodes that correspond to the billing dates.
        xmlNodeList = xmlDocument.SelectNodes("/nlBillingDates/nlBillingDate")

        'Size array to number of billing periods, which is equal to number of billing dates minus one.
        ReDim BillingPeriod(0 To xmlNodeList.Count - 1 - 1)

        'For each billing date minus the first, create a billing period in the array.
        For Each xmlNode In xmlNodeList
            If intNumberOfXmlNodes >= 1 Then
                BillingPeriod(intNumberOfXmlNodes - 1).PeriodOrder = xmlNode.Attributes("position").Value
                BillingPeriod(intNumberOfXmlNodes - 1).BillingDate = CDate(xmlNode.Attributes("date").Value)
                BillingPeriod(intNumberOfXmlNodes - 1).StartDate = datPeriodStartDate
                BillingPeriod(intNumberOfXmlNodes - 1).PeriodName = xmlNode.Attributes("label").Value
            End If
            datPeriodStartDate = DateAdd(DateInterval.Day, 1, CDate(xmlNode.Attributes("date").Value))
            intNumberOfXmlNodes = intNumberOfXmlNodes + 1
        Next

        Return BillingPeriod
    End Function


    ''' <summary>
    ''' Creates a a ribbon dropdown item.
    ''' </summary>
    ''' <returns>A ribbon dropdown item.</returns>
    ''' <remarks></remarks>
    Private Function CreateRibbonDropDownItem() As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem
        Return Me.Factory.CreateRibbonDropDownItem()
    End Function

    ''' <summary>
    ''' Generates the dropdown list with all billing periods.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GenerateBillingPeriodsDropDownList()
        Dim nlBillingPeriod As nlBillingPeriod
        NeolaneMonthDropDown.Items.Clear()
        For Each nlBillingPeriod In GetBillingPeriodsFromXML()
            Dim nlDropDownItem As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = CreateRibbonDropDownItem()
            nlDropDownItem.Label = nlBillingPeriod.PeriodName
            nlDropDownItem.ScreenTip = "Billing betwwen " & nlBillingPeriod.StartDate & " and " & nlBillingPeriod.BillingDate & "."
            NeolaneMonthDropDown.Items.Add(nlDropDownItem)
            If nlBillingPeriod.StartDate <= Now() And nlBillingPeriod.BillingDate >= Now() Then
                NeolaneMonthDropDown.SelectedItem = nlDropDownItem
            End If
        Next
    End Sub

    ''' <summary>
    ''' Actions to run when loading ribbon.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NeolaneRibbon_Load(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs) Handles Me.Load
        GenerateBillingPeriodsDropDownList()
    End Sub

    Private Function GetBillingPeriod(ByVal strInputBillingPeriodName As String) As nlBillingPeriod
        Dim nlBillingPeriod As nlBillingPeriod
        Dim nlAllPeriods As nlBillingPeriod()
        Dim intPeriodCount As Integer = 0

        nlAllPeriods = GetBillingPeriodsFromXML()

        Do Until nlAllPeriods(intPeriodCount).PeriodName = strInputBillingPeriodName Or intPeriodCount > nlAllPeriods.Length
            intPeriodCount = intPeriodCount + 1
        Loop

        If intPeriodCount <= nlAllPeriods.Length Then
            nlBillingPeriod = nlAllPeriods(intPeriodCount)
        Else
            nlBillingPeriod = Nothing
        End If

        Return nlBillingPeriod
    End Function

    ''' <summary>
    ''' Gets the billing period that surrounds a given date.
    ''' </summary>
    ''' <param name="datInputDate">Date from which lookup the billing period.</param>
    ''' <returns>A billing period.</returns>
    ''' <remarks></remarks>
    Private Function GetBillingPeriod(ByVal datInputDate As Date) As nlBillingPeriod
        Dim nlBillingPeriod As nlBillingPeriod
        Dim nlAllPeriods As nlBillingPeriod()
        Dim intPeriodCount As Integer = 0

        nlAllPeriods = GetBillingPeriodsFromXML()


        Do Until (nlAllPeriods(intPeriodCount).StartDate <= datInputDate And nlAllPeriods(intPeriodCount).BillingDate >= datInputDate) Or intPeriodCount > nlAllPeriods.Length
            intPeriodCount = intPeriodCount + 1
        Loop

        If intPeriodCount <= nlAllPeriods.Length Then
            nlBillingPeriod = nlAllPeriods(intPeriodCount)
        Else
            nlBillingPeriod = Nothing
        End If

        Return nlBillingPeriod
    End Function


    ''' <summary>
    ''' Launches the export of task list within period defined by current date.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Not working, fixed dates instead.</remarks>
    Private Sub LaunchBillingPeriodCurrentMonthButton_Click(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles LaunchBillingPeriodCurrentMonthButton.Click
        Dim nlBillingPeriod As nlBillingPeriod = GetBillingPeriod(Now())
        nlBillingFunctions.ListTasks(nlBillingPeriod.StartDate, nlBillingPeriod.BillingDate)
    End Sub



    ''' <summary>
    ''' Launches the export of task list within period defined by manual input.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub LaunchBillingPeriodCustomPeriodButton_Click(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles LaunchBillingPeriodCustomPeriodButton.Click
        Dim strInputStartDate As String = vbNull
        Dim strInputEndDate As String = vbNull

        'String formats http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.strings.format.aspx
        strInputStartDate = InputBox(Prompt:="Please input a START date for the period. All items on or after this date will be included." & Chr(13) & "Use format 'dd/mm/yyyy'." & Chr(13) & "Default is one month ago.", DefaultResponse:=Format(DateAdd(DateInterval.Day, 1, DateAdd(DateInterval.Month, -1, Today())), "dd/MM/yyyy"), Title:="Input start date")
        strInputEndDate = InputBox(Prompt:="Please input an END date for the period. All items on or before this date will be included." & Chr(13) & "Use format 'dd/mm/yyyy'." & Chr(13) & "Default is today.", DefaultResponse:=Format(Today(), "dd/MM/yyyy"), Title:="Input end date")

        If IsDate(strInputStartDate) And IsDate(strInputEndDate) Then
            nlBillingFunctions.ListTasks(strInputStartDate, strInputEndDate)
        Else
            MsgBox("Dates are invalid. Please check format.", MsgBoxStyle.Critical)
        End If
    End Sub

    ''' <summary>
    ''' Launches the export of task list within period defined by selection in NeolaneMonthDropDown.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub LaunchBillingPeriodSelectedMonthButton_Click(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles LaunchBillingPeriodSelectedMonthButton.Click
        Dim nlBillingPeriod As nlBillingPeriod = GetBillingPeriod(NeolaneMonthDropDown.SelectedItem.Label)
        nlBillingFunctions.ListTasks(nlBillingPeriod.StartDate, nlBillingPeriod.BillingDate)
    End Sub
End Class
