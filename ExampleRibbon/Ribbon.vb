Imports System.Runtime.InteropServices
Imports Extensibility
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls
Imports stdole

<ComVisible(True)>
<Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4")>
<ProgId("ExampleRibbon.Ribbon")>
Public Class Ribbon
    Implements IDTExtensibility2
    Implements IRibbonExtensibility

    Private ReadOnly _ribbon As Containers.Ribbon
    Private ReadOnly _extensions As ICollection(Of IDTExtensibility2)

    Public Sub New()
        _ribbon = BuildRibbon()
        _extensions = New IDTExtensibility2() { New Troubleshooter() }
    End Sub

    Private Function BuildRibbon() As Containers.Ribbon
        Dim tab As Tab = New TabBuilder().
                WithGroups(New ButtonsGroup(Me).AsGroup(), New DesktopFilesGroup(Me).AsGroup()).
                WithLabel("Example Tab").
                Build()

        Return New RibbonBuilder().
            OnLoad(AddressOf OnLoad).
            WithTab(tab).
            Build()
    End Function

    Public Function GetCustomUI(RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return _ribbon.RibbonX
    End Function

#Region "Callbacks"

    Public Sub OnSelectionChanged(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
        _ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex = selectedIndex
    End Sub

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer
        Return _ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex
    End Function

    Public Function GetSelectedItemId(control As IRibbonControl) As String
        Return _ribbon.GetElement(Of ISelect)(control.Id).Selected.ID
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String
        Return _ribbon.GetDropdownItem(control.Id, index).ScreenTip
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String
        Return _ribbon.GetDropdownItem(control.Id, index).SuperTip
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String
        Return _ribbon.GetDropdownItem(control.Id, index).Label
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp
        Return _ribbon.GetDropdownItem(control.Id, index).Image
    End Function

    Public Function GetItemId(control As IRibbonControl, index As Integer) As String
        Return _ribbon.GetDropdownItem(control.Id, index).ID
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer
        Return _ribbon.GetElement(Function(c As Container(Of Item)) c.ID.Equals(control.Id)).Count
    End Function

    Public Sub OnLoad(ribbon As IRibbonUI)
        _ribbon.AssignRibbonUI(ribbon)
    End Sub

    Public Sub OnAction(control As IRibbonControl)
        _ribbon.GetElement(Of IOnAction)(control.Id).Execute()
    End Sub

    Public Function GetImage(control As IRibbonControl) As IPictureDisp
        Return _ribbon.GetElement(Of IImage)(control.Id).Image
    End Function

#End Region

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        For Each extension As IDTExtensibility2 In _extensions
            extension.OnConnection(Application, ConnectMode, AddInInst, custom)
        Next
    End Sub

    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        For Each extension As IDTExtensibility2 In _extensions
            extension.OnDisconnection(RemoveMode, custom)
        Next
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        For Each extension As IDTExtensibility2 In _extensions
            extension.OnAddInsUpdate(custom)
        Next
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        For Each extension As IDTExtensibility2 In _extensions
            extension.OnStartupComplete(custom)
        Next
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        For Each extension As IDTExtensibility2 In _extensions
            extension.OnBeginShutdown(custom)
        Next
    End Sub

End Class