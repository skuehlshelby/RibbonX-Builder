Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports NetOffice.OfficeApi
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports stdole

Public Class SyntaxTesting
    Implements RibbonFactory.ICreateCallbacks

    Dim _ribbon As Ribbon = New Ribbon(startFromScratch:= False, [nameSpace]:= "Testing")

    Private Sub BuilderTest()
        With New TabBuilder
            .WithLabel("dfhfds").Visible().Build()
        End With

        Dim MyButton As Button = New ButtonBuilder().Large().WithLabel("My Button").ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello!")).Build()
        MyButton.Size = ControlSize.normal
    End Sub

    Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
        _ribbon.GetItem(Of IOnAction)(control.Id).Execute()
    End Sub

    Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
        _ribbon.GetItem(Of IOnChange)(control.Id).Execute()
    End Sub

    Public Sub SelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.SelectionChange
        Throw New NotImplementedException()
    End Sub

    Public Sub ButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.ButtonToggle
        _ribbon.GetItem(Of IPressed)(control.Id).Pressed = pressed
    End Sub

    Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
        Return _ribbon.GetItem(Of IEnabled)(control.Id).Enabled
    End Function

    Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
        Throw New NotImplementedException()
    End Function

    Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
        Throw New NotImplementedException()
    End Function

    Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
        Throw New NotImplementedException()
    End Function

    Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
        Throw New NotImplementedException()
    End Function

    Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
        Throw New NotImplementedException()
    End Function

    Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
        Throw New NotImplementedException()
    End Function

    Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
        Throw New NotImplementedException()
    End Function

    Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
        Throw New NotImplementedException()
    End Function

    Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
        Throw New NotImplementedException()
    End Function

    Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
        Throw New NotImplementedException()
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
        Return _ribbon.GetItem(Of ICollection(Of DropdownItem))(control.Id).Count
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
        Return _ribbon.GetItem(Of ICollection(Of DropdownItem))(control.Id).ElementAt(index).ID
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
        Throw New NotImplementedException()
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
        Throw New NotImplementedException()
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
        Throw New NotImplementedException()
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
        Throw New NotImplementedException()
    End Function

    Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
        Throw New NotImplementedException()
    End Function

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
        Throw New NotImplementedException()
    End Function
End Class