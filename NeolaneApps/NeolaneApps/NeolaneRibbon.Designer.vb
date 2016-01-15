Partial Class NeolaneRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.NeolaneTab = Me.Factory.CreateRibbonTab
        Me.BillingGroup = Me.Factory.CreateRibbonGroup
        Me.NeolaneMonthDropDown = Me.Factory.CreateRibbonDropDown
        Me.NeolaneBillingGallery = Me.Factory.CreateRibbonGallery
        Me.LaunchBillingPeriodCurrentMonthButton = Me.Factory.CreateRibbonButton
        Me.LaunchBillingPeriodSelectedMonthButton = Me.Factory.CreateRibbonButton
        Me.LaunchBillingPeriodCustomPeriodButton = Me.Factory.CreateRibbonButton
        Me.LaunchBillingPeriodCurrentWeekButton = Me.Factory.CreateRibbonButton
        Me.NeolaneTab.SuspendLayout()
        Me.BillingGroup.SuspendLayout()
        '
        'NeolaneTab
        '
        Me.NeolaneTab.Groups.Add(Me.BillingGroup)
        Me.NeolaneTab.Label = "Neolane"
        Me.NeolaneTab.Name = "NeolaneTab"
        '
        'BillingGroup
        '
        Me.BillingGroup.Items.Add(Me.NeolaneMonthDropDown)
        Me.BillingGroup.Items.Add(Me.NeolaneBillingGallery)
        Me.BillingGroup.Label = "Billing"
        Me.BillingGroup.Name = "BillingGroup"
        '
        'NeolaneMonthDropDown
        '
        Me.NeolaneMonthDropDown.Label = "Select Month"
        Me.NeolaneMonthDropDown.Name = "NeolaneMonthDropDown"
        '
        'NeolaneBillingGallery
        '
        Me.NeolaneBillingGallery.Buttons.Add(Me.LaunchBillingPeriodCurrentMonthButton)
        Me.NeolaneBillingGallery.Buttons.Add(Me.LaunchBillingPeriodSelectedMonthButton)
        Me.NeolaneBillingGallery.Buttons.Add(Me.LaunchBillingPeriodCustomPeriodButton)
        Me.NeolaneBillingGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.NeolaneBillingGallery.Label = "Launch billing..."
        Me.NeolaneBillingGallery.Name = "NeolaneBillingGallery"
        Me.NeolaneBillingGallery.OfficeImageId = "DollarSign"
        Me.NeolaneBillingGallery.ShowImage = True
        '
        'LaunchBillingPeriodCurrentMonthButton
        '
        Me.LaunchBillingPeriodCurrentMonthButton.Label = "... for current month."
        Me.LaunchBillingPeriodCurrentMonthButton.Name = "LaunchBillingPeriodCurrentMonthButton"
        '
        'LaunchBillingPeriodSelectedMonthButton
        '
        Me.LaunchBillingPeriodSelectedMonthButton.Label = "... for selected month."
        Me.LaunchBillingPeriodSelectedMonthButton.Name = "LaunchBillingPeriodSelectedMonthButton"
        '
        'LaunchBillingPeriodCustomPeriodButton
        '
        Me.LaunchBillingPeriodCustomPeriodButton.Label = "... for a custom period."
        Me.LaunchBillingPeriodCustomPeriodButton.Name = "LaunchBillingPeriodCustomPeriodButton"
        '
        'LaunchBillingPeriodCurrentWeekButton
        '
        Me.LaunchBillingPeriodCurrentWeekButton.Label = "... for selected month."
        Me.LaunchBillingPeriodCurrentWeekButton.Name = "LaunchBillingPeriodCurrentWeekButton"
        '
        'NeolaneRibbon
        '
        Me.Name = "NeolaneRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.NeolaneTab)
        Me.NeolaneTab.ResumeLayout(False)
        Me.NeolaneTab.PerformLayout()
        Me.BillingGroup.ResumeLayout(False)
        Me.BillingGroup.PerformLayout()

    End Sub

    Friend WithEvents NeolaneTab As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents BillingGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents NeolaneBillingGallery As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents LaunchBillingPeriodCurrentMonthButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents LaunchBillingPeriodSelectedMonthButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents LaunchBillingPeriodCustomPeriodButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents NeolaneMonthDropDown As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents LaunchBillingPeriodCurrentWeekButton As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As NeolaneRibbon
        Get
            Return Me.GetRibbon(Of NeolaneRibbon)()
        End Get
    End Property
End Class
