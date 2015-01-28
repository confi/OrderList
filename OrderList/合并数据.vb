Imports Microsoft.Office.Tools.Ribbon


Public Class 合并数据

    Private Sub 合并数据_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub count_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles countPID.Click
        Globals.ThisAddIn.makeOrderList()
    End Sub

    
    Private Sub countPipe_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles countPipe.Click
        Globals.ThisAddIn.makePipeList()
    End Sub
End Class
