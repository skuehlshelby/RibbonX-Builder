Namespace RibbonAttributes.Categories.SuperTip
    Friend NotInheritable Class GetSuperTip
        Inherits Supertip

        Private ReadOnly _callback As String

        Public Sub New(value As String, callback As FromControl(Of String))
            MyBase.New(value)
            Me._callback = callback.Method.Name
        End Sub

        Public Sub SetValue(value As String)
                Me.Value = value
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetSuperTip), _callback)
            End Get
        End Property
    End Class
End Namespace