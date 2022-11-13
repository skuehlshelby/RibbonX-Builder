Imports System.Text.RegularExpressions
Imports RibbonX.Controls.Base

Namespace Controls

    Public NotInheritable Class DialogLauncher
        Inherits RibbonElement

        Private ReadOnly _button As IButton

        Public Sub New(button As IButton)
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

                Return String.Join(Environment.NewLine, "<dialogBoxLauncher>", regex.Replace(_button.ConvertToXml(), String.Empty), "</dialogBoxLauncher>")
            End Get
        End Property

    End Class

End Namespace