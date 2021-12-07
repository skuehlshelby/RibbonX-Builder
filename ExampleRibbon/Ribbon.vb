Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.ExceptionServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports Extensibility
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports stdole

<
    ComVisible(True),
    Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4"),
    ProgId("ExampleRibbon.Ribbon")
>
Public Class Ribbon
    Implements IDTExtensibility2
    Implements IRibbonExtensibility

    Private ReadOnly _ribbon As Containers.Ribbon
    Private ReadOnly _displayDebugInfo As Boolean = False

    Public Sub New()
        If _displayDebugInfo Then
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf OnUnhandledException
            AddHandler AppDomain.CurrentDomain.FirstChanceException, AddressOf OnExceptionThrown
        End If

        _ribbon = BuildRibbon()
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

        Dim tab As Tab = New TabBuilder().
                WithGroup(group).
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

    Public Sub OnLoad(ribbon As IRibbonUI)
        _ribbon.AssignRibbonUI(ribbon)
    End Sub

    Public Sub OnAction(control As IRibbonControl)
        _ribbon.GetElement(Of IOnAction)(control.Id).Execute()
    End Sub

    Public Function GetImage(control As IRibbonControl) As IPictureDisp
        Return _ribbon.GetElement(Of IImage)(control.Id).Image
    End Function

#Region "Troubleshooting"

    Private Sub OnExceptionThrown(sender As Object, e As FirstChanceExceptionEventArgs)
        MessageBox(FormatException(e.Exception))
    End Sub

    Private Sub OnUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        MessageBox(FormatException(e.ExceptionObject))
    End Sub

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        If _displayDebugInfo Then MessageBox("Connected!")
    End Sub

    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        If _displayDebugInfo Then MessageBox("Disconnected!")
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        If _displayDebugInfo Then MessageBox("Add-Ins Updated!")
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        If _displayDebugInfo Then MessageBox("Startup Complete!")
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        If _displayDebugInfo Then MessageBox("Shutdown Starting!")
    End Sub

    Private Sub MessageBox(prompt As String)
        MsgBox(prompt, MsgBoxStyle.OkOnly, My.Application.Info.Title)
    End Sub

    Private Function FormatException(e As Exception) As String
        With New StringBuilder()
            .AppendLine($"A '{e.GetType().Name}' was thrown in method '{e.TargetSite.Name}'.")
            .AppendLine()
            .AppendLine("---------Message----------")
            .AppendLine(e.Message)
            .AppendLine()
            .AppendLine("-------Stack Trace--------")
            .AppendLine(e.StackTrace)

            If e.Data IsNot Nothing AndAlso e.Data.Count > 0 Then
                .AppendLine()
                .AppendLine("----------Data------------")
                For Each datum As Object In e.Data
                    .AppendLine(datum.ToString())
                Next
            End If

            Return .ToString()
        End With
    End Function
#End Region

End Class