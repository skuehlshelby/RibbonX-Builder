Imports RibbonX
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports Rx = RibbonX.RibbonXBuilder

''' <summary>
''' Regression coverage for a few low-severity edge cases:
'''  - invalidating the ribbon before OnLoad should raise the friendly
'''    MissingOnLoadException rather than a raw NullReferenceException;
'''  - setting an edit box / combo box Text to Nothing should raise a clear
'''    ArgumentNullException rather than a raw NullReferenceException.
''' </summary>
<TestClass()>
Public Class RegressionTests

    Private Function GetText(control As IRibbonControl) As String
        Return String.Empty
    End Function

    Private Sub OnChange(control As IRibbonControl, text As String)
    End Sub

    Private Shared Function BuildUnloadedRibbon() As IRibbon
        Dim group As IGroup = Rx.Group(Sub(g) g.WithControls(Rx.Button()))
        Dim tab As ITab = Rx.Tab(Sub(t) t.WithGroups(group))
        Return Rx.Ribbon(Sub(r) r.WithTabs(tab))
    End Function

    <TestMethod>
    Public Sub InvalidateControl_BeforeOnLoad_ThrowsMissingOnLoad()
        Dim ribbon As IRibbon = BuildUnloadedRibbon()

        Assert.ThrowsException(Of MissingOnLoadException)(Sub() ribbon.InvalidateControl("any"))
    End Sub

    <TestMethod>
    Public Sub Invalidate_BeforeOnLoad_ThrowsMissingOnLoad()
        Dim ribbon As IRibbon = BuildUnloadedRibbon()

        Assert.ThrowsException(Of MissingOnLoadException)(Sub() ribbon.Invalidate())
    End Sub

    <TestMethod>
    Public Sub ComboBoxText_SetToNothing_ThrowsArgumentNull()
        Dim combo As IComboBox = Rx.ComboBox(Sub(b) b.WithText("initial", AddressOf GetText, AddressOf OnChange))

        Assert.ThrowsException(Of ArgumentNullException)(Sub() combo.Text = Nothing)
    End Sub

    <TestMethod>
    Public Sub EditBoxText_SetToNothing_ThrowsArgumentNull()
        Dim edit As IEditBox = Rx.EditBox(Sub(b) b.WithText("initial", AddressOf GetText, AddressOf OnChange))

        Assert.ThrowsException(Of ArgumentNullException)(Sub() edit.Text = Nothing)
    End Sub

End Class
