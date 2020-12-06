Imports IRibbonControl = NetOffice.OfficeApi.IRibbonControl
Imports IPictureDisp = stdole.IPictureDisp

Namespace CallbackSignatures
    Interface ICreateCallbacks
        Function GetEnabled(ByVal control As IRibbonControl) As Boolean
        Function GetVisible(ByVal control As IRibbonControl) As Boolean
        Function GetLabel(ByVal control As IRibbonControl) As String
        Function GetShowLabel(ByVal control As IRibbonControl) As Boolean
        Function GetDescription(ByVal control As IRibbonControl) As String
        Function GetScreenTip(ByVal control As IRibbonControl) As String
        Function GetSuperTip(ByVal control As IRibbonControl) As String
        Function GetSize(ByVal control As IRibbonControl) As ControlSize
        Function GetShowImage(ByVal control As IRibbonControl) As Boolean
        Function GetPressed(ByVal control As IRibbonControl) As Boolean
        Function GetText(ByVal control As IRibbonControl) As String
        Function GetItemCount(ByVal control As IRibbonControl) As Integer
        Function GetItemID(ByVal control As IRibbonControl, index As Integer) As String
        Function GetItemImage(ByVal control As IRibbonControl, index As Integer) As IPictureDisp
        Function GetItemLabel(ByVal control As IRibbonControl, index As Integer) As String
        Function GetItemScreentip(ByVal control As IRibbonControl, index As Integer) As String
        Function GetItemSuperTip(ByVal control As IRibbonControl, index As Integer) As String
        Function GetSelectedItemID(ByVal control As IRibbonControl) As String
        Function GetSelectedItemIndex(ByVal control As IRibbonControl) As Integer
        Sub OnAction(ByVal control As IRibbonControl)
        Sub OnChange(ByVal control As IRibbonControl, text As String)
        Sub SelectionChange(ByVal control As IRibbonControl, selectedId As String, selectedIndex As Integer)
        Sub ButtonToggle(ByVal control As IRibbonControl, pressed As Boolean)
    End Interface
End Namespace
