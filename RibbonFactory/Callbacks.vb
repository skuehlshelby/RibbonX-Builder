Imports IRibbonControl = NetOffice.OfficeApi.IRibbonControl

Namespace CallbackSignatures

    Public Delegate Function FromControl(Of T)(ByVal Control As IRibbonControl) As T

    Public Delegate Function FromControlAndIndex(Of T)(ByVal Control As IRibbonControl, ByVal Index As Integer) As T

    Public Delegate Sub SelectionChanged(ByVal control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Public Delegate Sub ButtonPressed(ByVal control As IRibbonControl, ByVal Pressed As Boolean)

    Public Delegate Sub TextChanged(ByVal control As IRibbonControl, ByVal Text As String)

    Public Delegate Sub OnAction(ByVal control As IRibbonControl)

End Namespace

