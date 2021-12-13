Imports System.Drawing
Imports System.IO
Imports System.Reflection
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

        Dim piButton As Button = New ButtonBuilder().
                WithLabel("Calculate Pi", copyToScreenTip:= True).
                WithSuperTip("So that you too can know the value of Pi.").
                WithImage(New Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExampleRibbon.pi.png")), AddressOf GetImage).
                ThatDoes(AddressOf OnAction, Sub() MessageBox($"The value of Pi is {Math.PI:#.#####}...")).
                Build()

        Dim helloButton As Button = New ButtonBuilder().
                WithLabel("Hello World", copyToScreenTip:= True).
                WithSuperTip("The classic introductory exercise. Click me!").
                WithImage(New Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExampleRibbon.hello.png")), AddressOf GetImage).
                ThatDoes(AddressOf OnAction, Sub() MessageBox("Hello World!")).
                Build()

        Dim group As Group = New GroupBuilder().
                WithControls(helloButton, piButton).
                WithLabel("Example Group").
                Build()

        Dim fileDropDown As Dropdown = New DropdownBuilder().
                WithScreenTip("Desktop Files").
                WithSuperTip("The files on your desktop.").
                WithSize("A DropDown This Big").
                GetItemCountFrom(AddressOf GetItemCount).
                GetItemIdFrom(AddressOf GetItemId).
                GetItemImageFrom(AddressOf GetItemImage).
                GetItemLabelFrom(AddressOf GetItemLabel).
                GetItemSuperTipFrom(AddressOf GetItemSuperTip).
                GetItemScreenTipFrom(AddressOf GetItemScreenTip).
                GetSelectedItemIdFrom(AddressOf GetSelectedItemId).
                ThatDoes(Sub() Return, AddressOf OnSelectionChanged).
                Build()

        Dim dropDownLabel As LabelControl = New LabelControlBuilder().
                WithLabel("Desktop Files:").
                ShowLabel().
                Build()

        Dim launchButton As Button = New ButtonBuilder().
                WithLabel("Open File").
                WithSuperTip("Open or launch the selected file/program.").
                WithImage(Enums.ImageMSO.Common.FileOpen).
                ThatDoes(AddressOf OnAction, Sub() OpenFile(fileDropDown.Selected.SuperTip)).
                Build()

        Dim dropDownGroup As Group = New GroupBuilder().
                WithControl(BoxBuilder.Horizontal(launchButton, BoxBuilder.Vertical(dropDownLabel, fileDropDown))).
                WithLabel("Desktop Files").
                Build()

        Dim tab As Tab = New TabBuilder().
                WithGroups(group, dropDownGroup).
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

    Private Function GetSelectedItemIndex(control As IRibbonControl) As Integer
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
        _ribbon.SuspendLayout()

        With _ribbon.GetElement(Function(d As DropDown) True)
            For Each file As FileInfo In GetFilesOnDesktop()
                .Add(ConvertFileToDropDownItem(file, AddressOf GetImage))
            Next
        End With

        _ribbon.ResumeLayout()

        For Each extension In _extensions
            extension.OnConnection(Application, ConnectMode, AddInInst, custom)
        Next
    End Sub

    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        For Each extension In _extensions
            extension.OnDisconnection(RemoveMode, custom)
        Next
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        For Each extension In _extensions
            extension.OnAddInsUpdate(custom)
        Next
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        For Each extension In _extensions
            extension.OnStartupComplete(custom)
        Next
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        For Each extension In _extensions
            extension.OnBeginShutdown(custom)
        Next
    End Sub

    Private Sub MessageBox(prompt As String)
        MsgBox(prompt, MsgBoxStyle.OkOnly, My.Application.Info.Title)
    End Sub

End Class