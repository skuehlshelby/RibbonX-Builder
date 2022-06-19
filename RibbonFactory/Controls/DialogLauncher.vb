Imports System.Text.RegularExpressions
Imports RibbonX.Controls.Base

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
                Dim regex As Regex = New Regex("(?:size|getSize|visible|getVisible)=""\w+""")

                Return String.Join(Environment.NewLine, "<dialogBoxLauncher>", regex.Replace(_button.XML, String.Empty), "</dialogBoxLauncher>")
            End Get
        End Property

    End Class

End Namespace