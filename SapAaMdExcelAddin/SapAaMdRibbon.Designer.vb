Partial Class SapAaMdRibbon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapAaMdRibbon))
        Me.SapAaMd = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ButtonSapAssetCreate = Me.Factory.CreateRibbonButton
        Me.ButtonSapAssetlChange = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ButtonSapLegValCheck = Me.Factory.CreateRibbonButton
        Me.ButtonSapLegValPost = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapAaMd.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapAaMd
        '
        Me.SapAaMd.Groups.Add(Me.Group1)
        Me.SapAaMd.Groups.Add(Me.Group3)
        Me.SapAaMd.Groups.Add(Me.Group2)
        Me.SapAaMd.Label = "SAP AA Md"
        Me.SapAaMd.Name = "SapAaMd"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonSapAssetCreate)
        Me.Group1.Items.Add(Me.ButtonSapAssetlChange)
        Me.Group1.Label = "Asset Master"
        Me.Group1.Name = "Group1"
        '
        'ButtonSapAssetCreate
        '
        Me.ButtonSapAssetCreate.Image = CType(resources.GetObject("ButtonSapAssetCreate.Image"), System.Drawing.Image)
        Me.ButtonSapAssetCreate.Label = "Create Asset"
        Me.ButtonSapAssetCreate.Name = "ButtonSapAssetCreate"
        Me.ButtonSapAssetCreate.ShowImage = True
        '
        'ButtonSapAssetlChange
        '
        Me.ButtonSapAssetlChange.Image = CType(resources.GetObject("ButtonSapAssetlChange.Image"), System.Drawing.Image)
        Me.ButtonSapAssetlChange.Label = "Change Asset"
        Me.ButtonSapAssetlChange.Name = "ButtonSapAssetlChange"
        Me.ButtonSapAssetlChange.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ButtonSapLegValCheck)
        Me.Group3.Items.Add(Me.ButtonSapLegValPost)
        Me.Group3.Label = "Legacy Values"
        Me.Group3.Name = "Group3"
        '
        'ButtonSapLegValCheck
        '
        Me.ButtonSapLegValCheck.Image = CType(resources.GetObject("ButtonSapLegValCheck.Image"), System.Drawing.Image)
        Me.ButtonSapLegValCheck.Label = "Legacy Values Check"
        Me.ButtonSapLegValCheck.Name = "ButtonSapLegValCheck"
        Me.ButtonSapLegValCheck.ShowImage = True
        '
        'ButtonSapLegValPost
        '
        Me.ButtonSapLegValPost.Image = CType(resources.GetObject("ButtonSapLegValPost.Image"), System.Drawing.Image)
        Me.ButtonSapLegValPost.Label = "Legacy Values Post"
        Me.ButtonSapLegValPost.Name = "ButtonSapLegValPost"
        Me.ButtonSapLegValPost.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButtonLogon)
        Me.Group2.Items.Add(Me.ButtonLogoff)
        Me.Group2.Label = "SAP Logon"
        Me.Group2.Name = "Group2"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'SapAaMdRibbon
        '
        Me.Name = "SapAaMdRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapAaMd)
        Me.SapAaMd.ResumeLayout(False)
        Me.SapAaMd.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapAaMd As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapAssetCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapAssetlChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSapLegValCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSapLegValPost As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapAaMdRibbon
        Get
            Return Me.GetRibbon(Of SapAaMdRibbon)()
        End Get
    End Property
End Class
