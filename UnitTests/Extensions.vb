Imports System.Runtime.CompilerServices
Imports RibbonFactory

Friend Module Extensions

    <Extension()>
    Public Sub ValueChangedIsRaised(Of T As RibbonElement)(assert As Assert, element As T, action As Action(Of T))
        Dim timesRaised As Integer = 0
        Dim handler As EventHandler(Of ValueChangedEventArgs) = New EventHandler(Of ValueChangedEventArgs)(Sub(o, e) timesRaised += 1)

        AddHandler element.ValueChanged, handler

        action.Invoke(element)

        Assert.AreEqual(1, timesRaised)
    End Sub

End Module
