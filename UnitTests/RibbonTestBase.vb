Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Public Class RibbonTestBase
    Implements IRibbonExtensibility
    Implements ICreateCallbacks

    Protected Ribbon As Ribbon

    Public Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return Ribbon.RibbonX
    End Function

    Protected Sub AssignFakeIRibbonUI()
        Ribbon.AssignRibbonUI(New Stubs.Excel(Me))
    End Sub

    Protected Function MakeButton() As Button
        Return New ButtonBuilder() _
                .WithLabel("My Button") _
                .WithSuperTip("Super") _
                .WithImage(Common.DollarSign) _
                .Disabled(AddressOf GetEnabled) _
                .Build()
    End Function

    Protected Function GetMenu() As Menu
        Return New MenuBuilder() _
            .WithLabel("My Menu").InsertBeforeMSO(Excel.Calculator).WithControl(MakeButton()).WithLargeItems().
            ShowLabel().WithImage(Common.SadFace).Build()
    End Function

    Protected Function MakeRibbonWithOneTabAndOneGroup(ParamArray controls As RibbonElement()) As Ribbon
        Return New RibbonBuilder() _
            .WithTab(New TabBuilder().WithLabel("Test Tab") _
                        .WithGroup(New GroupBuilder().WithLabel("Test Group") _
                                      .WithControls(controls).Build()).Build()).Build()
    End Function

    Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
        Return Ribbon.GetElement(Of IEnabled)(control.Id).Enabled
    End Function

    Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
        Return Ribbon.GetElement(Of IVisible)(control.Id).Visible
    End Function

    Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
        Return Ribbon.GetElement(Of ILabel)(control.Id).Label
    End Function

    Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
        Return Ribbon.GetElement(Of IShowLabel)(control.Id).ShowLabel
    End Function

    Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
        Return Ribbon.GetElement(Of IDescription)(control.Id).Description
    End Function

    Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
        Return Ribbon.GetElement(Of IScreenTip)(control.Id).ScreenTip
    End Function

    Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
        Return Ribbon.GetElement(Of ISuperTip)(control.Id).SuperTip
    End Function

    Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
        Return Ribbon.GetElement(Of ISize)(control.Id).Size
    End Function

    Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
        Return Ribbon.GetElement(Of IShowImage)(control.Id).ShowImage
    End Function

    Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
        Return Ribbon.GetElement(Of IPressed)(control.Id).Pressed
    End Function

    Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
        Return Ribbon.GetElement(Of IText)(control.Id).Text
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
        Return Ribbon.GetElement(Of IEnumerable(Of RibbonElement))(control.Id).Count()
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
        Return Ribbon.GetDropdownItem(control.Id, index).ID
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
        Return Ribbon.GetDropdownItem(control.Id, index).Image
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
        Return Ribbon.GetDropdownItem(control.Id, index).Label
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
        Return Ribbon.GetDropdownItem(control.Id, index).ScreenTip
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
        Return Ribbon.GetDropdownItem(control.Id, index).SuperTip
    End Function

    Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
        Return Ribbon.GetElement(Of ISelect)(control.Id).Selected.ID
    End Function

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
        Return Ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex
    End Function

    Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
        Ribbon.GetElement(Of IOnAction)(control.Id).Execute()
    End Sub

    Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
        Ribbon.GetElement(Of IText)(control.Id).Text = text
        Ribbon.GetElement(Of IOnChange)(control.Id).Execute()
    End Sub

    Public Sub SelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.SelectionChange
        Ribbon.GetElement(Of ISelect)(control.Id).Selected = Ribbon.GetDropdownItem(control.Id, selectedIndex)
        Ribbon.GetElement(Of IOnAction)(control.Id).Execute()
    End Sub

    Public Sub ButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.ButtonToggle
        With Ribbon.GetElement(Of ToggleButton)(control.Id)
            .Pressed = pressed

            If pressed Then
                .Execute()
            End If
        End With
    End Sub

    Public Function GetItemHeight(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemHeight
        Throw New NotImplementedException()
    End Function

    Public Function GetItemWidth(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemWidth
        Throw New NotImplementedException()
    End Function

    Public Sub OnLoad(ribbon As IRibbonUI) Implements ICreateCallbacks.OnLoad
        Throw New NotImplementedException()
    End Sub

    Public Function LoadImage(imageID As String) As IPictureDisp Implements ICreateCallbacks.LoadImage
        Throw New NotImplementedException()
    End Function
End Class