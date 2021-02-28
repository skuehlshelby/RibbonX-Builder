Namespace RibbonAttributes.Categories.Label
    Friend Class ShowLabel
        Inherits RibbonAttribute(Of Boolean, ShowLabel)

        Public Sub New(Value As Boolean)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(ShowLabel)), CamelCase(GetValue()))
            End Get
        End Property
    End Class
End Namespace
