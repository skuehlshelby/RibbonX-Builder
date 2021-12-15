Namespace Controls
    Public NotInheritable Class DialogLauncher
        Inherits RibbonElement

        Private ReadOnly _button As Button

        Public Sub New(button As Button)
            _button = button
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return String.Empty
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Join(Environment.NewLine, "<dialogBoxLauncher>", _button.XML, "</dialogBoxLauncher>")
            End Get
        End Property

    End Class

End NameSpace