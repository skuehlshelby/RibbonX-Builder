Imports System.Reflection
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.SimpleTypes
Imports System.Drawing

Public MustInherit Class TestBase
    Implements ICreateCallbacks

    Public MustOverride Sub NullTemplate_NoThrow()

    Public MustOverride Sub NullConfiguration_NoThrow()

    Public MustOverride Sub ProducesLegalRibbonX()

    Public MustOverride Sub ContainsNoNullValuesByDefault()

    Public MustOverride Sub PropertiesAreMappedCorrectly()

    Public MustOverride Sub PropertiesWithoutCallbacksCannotBeModified()

    Public MustOverride Sub PropertiesWithCallbacksCanBeModified()

    Public MustOverride Sub TemplatePropertiesAreCopiedToNewControl()

    Protected Shared Function BlankBitmap() As Bitmap
        Return New Bitmap(24, 24)
    End Function

#Region "Non-Static Callback Methods"

    Public Sub OnLoad(ribbon As IRibbonUI) Implements ICreateCallbacks.OnLoad
        Throw New NotSupportedException()
    End Sub

    Public Function LoadImage(imageID As String) As IPictureDisp Implements ICreateCallbacks.LoadImage
        Throw New NotSupportedException()
    End Function

    Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
        Throw New NotSupportedException()
    End Function

    Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
        Throw New NotSupportedException()
    End Function

    Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
        Throw New NotSupportedException()
    End Function

    Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
        Throw New NotSupportedException()
    End Function

    Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
        Throw New NotSupportedException()
    End Function

    Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
        Throw New NotSupportedException()
    End Function

    Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
        Throw New NotSupportedException()
    End Function

    Public Function GetKeyTip(control As IRibbonControl) As KeyTip Implements ICreateCallbacks.GetKeyTip
        Throw New NotSupportedException()
    End Function

    Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
        Throw New NotSupportedException()
    End Function

    Public Function GetImage(control As IRibbonControl) As IPictureDisp Implements ICreateCallbacks.GetImage
        Throw New NotSupportedException()
    End Function

    Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
        Throw New NotSupportedException()
    End Function

    Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
        Throw New NotSupportedException()
    End Function

    Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
        Throw New NotSupportedException()
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
        Throw New NotSupportedException()
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
        Throw New NotSupportedException()
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
        Throw New NotSupportedException()
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
        Throw New NotSupportedException()
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
        Throw New NotSupportedException()
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
        Throw New NotSupportedException()
    End Function

    Public Function GetItemHeight(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemHeight
        Throw New NotSupportedException()
    End Function

    Public Function GetItemWidth(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemWidth
        Throw New NotSupportedException()
    End Function

    Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
        Throw New NotSupportedException()
    End Function

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
        Throw New NotSupportedException()
    End Function

    Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
        Throw New NotSupportedException()
    End Sub

    Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
        Throw New NotSupportedException()
    End Sub

    Public Sub OnSelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.OnSelectionChange
        Throw New NotSupportedException()
    End Sub

    Public Sub OnButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.OnButtonToggle
        Throw New NotSupportedException()
    End Sub

#End Region

    'The below callback methods would not actually work in a real add-in. They are only for testing purposes.

#Region "Static Callback Methods"

    Public Shared Sub OnLoadShared(ribbon As IRibbonUI)
        Throw New NotSupportedException()
    End Sub

    Public Shared Function LoadImageShared(imageID As String) As IPictureDisp
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetEnabledShared(control As IRibbonControl) As Boolean
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetVisibleShared(control As IRibbonControl) As Boolean
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetLabelShared(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetShowLabelShared(control As IRibbonControl) As Boolean
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetDescriptionShared(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetScreenTipShared(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetSuperTipShared(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetKeyTipShared(control As IRibbonControl) As KeyTip
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetSizeShared(control As IRibbonControl) As ControlSize
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetImageShared(control As IRibbonControl) As IPictureDisp
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetShowImageShared(control As IRibbonControl) As Boolean
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetPressedShared(control As IRibbonControl) As Boolean
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetTextShared(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemCountShared(control As IRibbonControl) As Integer
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemIDShared(control As IRibbonControl, index As Integer) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemImageShared(control As IRibbonControl, index As Integer) As IPictureDisp
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemLabelShared(control As IRibbonControl, index As Integer) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetTitle(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemScreenTipShared(control As IRibbonControl, index As Integer) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemSuperTipShared(control As IRibbonControl, index As Integer) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemHeightShared(control As IRibbonControl) As Integer
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetItemWidthShared(control As IRibbonControl) As Integer
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetSelectedItemIDShared(control As IRibbonControl) As String
        Throw New NotSupportedException()
    End Function

    Public Shared Function GetSelectedItemIndexShared(control As IRibbonControl) As Integer
        Throw New NotSupportedException()
    End Function

    Public Shared Sub OnActionShared(control As IRibbonControl)
        Throw New NotSupportedException()
    End Sub

    Public Shared Sub OnChangeShared(control As IRibbonControl, text As String)
        Throw New NotSupportedException()
    End Sub

    Public Shared Sub OnSelectionChangeShared(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
        Throw New NotSupportedException()
    End Sub

    Public Shared Sub OnButtonToggleShared(control As IRibbonControl, pressed As Boolean)
        Throw New NotSupportedException()
    End Sub

#End Region

End Class
