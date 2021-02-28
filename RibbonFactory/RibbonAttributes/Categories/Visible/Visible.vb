Namespace RibbonAttributes.Categories.Visible
    Friend Class Visible
        Inherits RibbonAttribute(Of Boolean, Visible)

        Public Sub New(Value As Boolean)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(Visible)), CamelCase(GetValue()))
            End Get
        End Property
    End Class
End Namespace