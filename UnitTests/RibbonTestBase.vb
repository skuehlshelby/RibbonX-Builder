Imports System.IO
Imports System.Text
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Enums
Imports stdole

Public Class RibbonTestBase
    Implements ICreateCallbacks

    Private _originalConsoleOutputStream As TextWriter
    Private ReadOnly _builder As StringBuilder = new StringBuilder()
    Private ReadOnly _redirectedConsoleOutput As StringWriter = new StringWriter(_builder)
    
    Protected Sub RedirectConsoleOutput()
        _originalConsoleOutputStream = Console.Out
        Console.SetOut(_redirectedConsoleOutput)
    End Sub
    
    Protected Sub RestoreOriginalConsoleOutput()
        Console.SetOut(_originalConsoleOutputStream)
    End Sub
    
    Protected Function GetContentsOfRedirectedConsoleOutput() As String
        Return _builder.ToString()
    End Function
    
    Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
        Throw New NotImplementedException
    End Function

    Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
        Throw New NotImplementedException
    End Function

    Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
        Throw New NotImplementedException
    End Function

    Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
        Throw New NotImplementedException
    End Function

    Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
        Throw New NotImplementedException
    End Function

    Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
        Throw New NotImplementedException
    End Function

    Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
        Throw New NotImplementedException
    End Function

    Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
        Throw New NotImplementedException
    End Function

    Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
        Throw New NotImplementedException
    End Function

    Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
        Throw New NotImplementedException
    End Function

    Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
        Throw New NotImplementedException
    End Function

    Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
        Throw New NotImplementedException
    End Function

    Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
        Throw New NotImplementedException
    End Function

    Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
        Throw New NotImplementedException
    End Function

    Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
        Throw New NotImplementedException
    End Function

    Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
        Throw New NotImplementedException
    End Function

    Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
        Throw New NotImplementedException
    End Function

    Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
        Throw New NotImplementedException
    End Function

    Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
        Throw New NotImplementedException
    End Function

    Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
        Throw New NotImplementedException
    End Sub

    Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
        Throw New NotImplementedException
    End Sub

    Public Sub SelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.SelectionChange
        Throw New NotImplementedException
    End Sub

    Public Sub ButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.ButtonToggle
        Throw New NotImplementedException
    End Sub
End Class