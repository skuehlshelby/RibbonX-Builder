Imports System.Runtime.InteropServices
Imports Extensibility
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers

<ComVisible(True)>
<Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4")>
<ProgId("ExampleRibbon.Ribbon")>
Public Class Ribbon
    Inherits Utilities.StockRibbonBase
    Implements IDTExtensibility2
    Implements IRibbonExtensibility

    Public Sub New()
        MyBase.New(New Troubleshooter())
    End Sub

    Protected Overrides Function BuildRibbon() As Containers.Ribbon
        Dim builtInGroupTest As Group = New GroupBuilder().WithIdMso(Enums.MSO.Excel.GroupFont).Build()

        Dim tab As Tab = New TabBuilder().
                WithGroups(New ButtonsGroup(Me).AsGroup(), New DesktopFilesGroup(Me).AsGroup(), builtInGroupTest).
                WithLabel("Example Tab").
                Build()

        Return New RibbonBuilder().
            OnLoad(AddressOf OnLoad).
            WithTab(tab).
            Build()
    End Function

End Class