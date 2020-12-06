Imports NUnit.Framework
Imports RibbonFactory
Public Class ButtonTests
    <Test>
    Public Sub ButtonIsCreatable()
        Dim MyButton As Button = New ButtonBuilder().WithLabel("The Label").WithSupertip("The Supertip").Build()

        System.Diagnostics.Debug.Write(MyButton.XML)
    End Sub
End Class
