Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums

<
    ComVisible(True),
    Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4"),
    ProgId("Example." & NameOf(Ribbon))
>
Public Class Ribbon
    Implements IRibbonExtensibility

    Private ReadOnly _ribbon As Containers.Ribbon = BuildRibbon()

    Private Function BuildRibbon() As Containers.Ribbon
        Dim button As Button = New ButtonBuilder().
                WithLabel("Example Button").
                WithSuperTip("This button is an example. Click me!").
                WithImage(ImageMSO.Common.TraceError).
                ThatDoes(AddressOf OnAction, Sub() MsgBox("Hello World!")).
                Build()

        Dim group As Group = New GroupBuilder().
                WithControl(button).
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

    Public Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return _ribbon.RibbonX
    End Function

    Public Sub OnLoad(ribbon As IRibbonUI)
        _ribbon.AssignRibbonUI(ribbon)
    End Sub

    Public Sub OnAction(control As IRibbonControl)
        _ribbon.GetElement(Of IOnAction)(control.Id).Execute()
    End Sub

End Class