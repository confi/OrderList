Partial Class 合并数据
    Inherits Microsoft.Office.Tools.Ribbon.OfficeRibbon

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = New Microsoft.Office.Tools.Ribbon.RibbonTab
        Me.Group1 = New Microsoft.Office.Tools.Ribbon.RibbonGroup
        Me.count = New Microsoft.Office.Tools.Ribbon.RibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "AWSTEC"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.count)
        Me.Group1.Label = "合并数据清单"
        Me.Group1.Name = "Group1"
        '
        'count
        '
        Me.count.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.count.Image = Global.OrderList.My.Resources.Resources.logo图
        Me.count.ImageName = "AWSlogo"
        Me.count.Label = "统计数量"
        Me.count.Name = "count"
        Me.count.ShowImage = True
        '
        '合并数据
        '
        Me.Name = "合并数据"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents count As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection
    Inherits Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property 合并数据() As 合并数据
        Get
            Return Me.GetRibbon(Of 合并数据)()
        End Get
    End Property
End Class
