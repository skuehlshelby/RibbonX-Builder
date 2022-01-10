Imports Microsoft.Office.Core
Imports RibbonFactory.Enums
Imports stdole

''' <summary>
''' This interface exists only for convenience.
''' Implementing it will cause all possible callback signatures to be generated in the implementing class.
''' This is useful for creating dynamic ribbon objects, since the objects require callback methods that already exist.
''' </summary>
Public Interface ICreateCallbacks
    Sub OnLoad(ribbon As IRibbonUI)
    Function LoadImage(imageID As String) As IPictureDisp
    Function GetEnabled(control As IRibbonControl) As Boolean
    Function GetVisible(control As IRibbonControl) As Boolean
    Function GetLabel(control As IRibbonControl) As String
    Function GetShowLabel(control As IRibbonControl) As Boolean
    Function GetDescription(control As IRibbonControl) As String
    Function GetScreenTip(control As IRibbonControl) As String
    Function GetSuperTip(control As IRibbonControl) As String
    Function GetKeyTip(control As IRibbonControl) As KeyTip
    Function GetSize(control As IRibbonControl) As ControlSize
    Function GetImage(control As IRibbonControl) As IPictureDisp
    Function GetShowImage(control As IRibbonControl) As Boolean
    Function GetPressed(control As IRibbonControl) As Boolean
    Function GetText(control As IRibbonControl) As String
    Function GetItemCount(control As IRibbonControl) As Integer
    Function GetItemID(control As IRibbonControl, index As Integer) As String
    Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp
    Function GetItemLabel(control As IRibbonControl, index As Integer) As String
    Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String
    Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String
    Function GetItemHeight(control As IRibbonControl) As Integer
    Function GetItemWidth(control As IRibbonControl) As Integer
    Function GetSelectedItemID(control As IRibbonControl) As String
    Function GetSelectedItemIndex(control As IRibbonControl) As Integer
    Sub OnAction(control As IRibbonControl)
    Sub OnChange(control As IRibbonControl, text As String)
    Sub OnSelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Sub OnButtonToggle(control As IRibbonControl, pressed As Boolean)
End Interface
