Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Extensibility
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports stdole

<
    ComVisible(True),
    Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4"),
    ProgId("ExampleRibbon." & NameOf(Ribbon))
>
Public Class Ribbon
    Implements IDTExtensibility2
    Implements IRibbonExtensibility

    Private ReadOnly _ribbon As Containers.Ribbon

    Public Sub New()
        _ribbon = BuildRibbon()
    End Sub

    Private Function BuildRibbon() As Containers.Ribbon
        Dim piButton As Button = New ButtonBuilder().
                WithLabel("Pi").
                WithSuperTip("What is Pi?").
                WithImage(ImageMSO.Misc.Chart3DPieChart).
                ThatDoes(AddressOf OnAction, Sub() MsgBox(Math.PI.ToString("#.##"), MsgBoxStyle.OkOnly, My.Application.Info.Title)).
                Build()

        Dim messageBoxButton As Button = New ButtonBuilder().
                WithLabel("Example Button").
                WithSuperTip("This button is an example. Click me!").
                WithImage(ImageMSO.Common.DollarSign).
                ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!", MsgBoxStyle.OkOnly, My.Application.Info.Title)).
                Build()

        Dim group As Group = New GroupBuilder().
                WithControls(messageBoxButton, piButton).
                WithLabel("Example Group").
                Build()

        Dim tab As Tab = New TabBuilder().
                WithGroup(group).
                WithLabel("Example Tab").
                Build()

        Dim ribbon As Containers.Ribbon = New RibbonBuilder().
            OnLoad(AddressOf OnLoad).
            WithTab(tab).
            Build()

        Dim errors As RibbonErrorLog = ribbon.GetErrors()

        Return ribbon
    End Function

    Public Function GetCustomUI(RibbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return _ribbon.RibbonX
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

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        MsgBox("Connected!")
    End Sub

    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        MsgBox("Disconnected!")
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        MsgBox("Add-Ins Updated!")
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        MsgBox("Startup Complete!")
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        MsgBox("Shutdown Starting!")
    End Sub
End Class