Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.ComTypes.StdOle

Namespace Callbacks

    Public Delegate Sub OnLoad(ribbon As IRibbonUI)

    Public Delegate Function LoadImage(imageID As String) As IPictureDisp

    Public Delegate Function FromControl(Of T)(control As IRibbonControl) As T

    Public Delegate Function FromControlAndIndex(Of T)(control As IRibbonControl, index As Integer) As T

    Public Delegate Sub SelectionChanged(control As IRibbonControl, selectedId As String, selectedIndex As Integer)

    Public Delegate Sub ButtonPressed(control As IRibbonControl, pressed As Boolean)

    Public Delegate Sub TextChanged(control As IRibbonControl, text As String)

    Public Delegate Sub OnAction(control As IRibbonControl)

End Namespace