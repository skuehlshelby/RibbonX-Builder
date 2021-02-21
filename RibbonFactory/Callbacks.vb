Imports NetOffice.OfficeApi


Public Delegate Function FromControl(Of T)(Control As IRibbonControl) As T

Public Delegate Function FromControlAndIndex(Of T)(Control As IRibbonControl, Index As Integer) As T

Public Delegate Sub SelectionChanged(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

Public Delegate Sub ButtonPressed(control As IRibbonControl, Pressed As Boolean)

Public Delegate Sub TextChanged(control As IRibbonControl, Text As String)

Public Delegate Sub OnAction(control As IRibbonControl)