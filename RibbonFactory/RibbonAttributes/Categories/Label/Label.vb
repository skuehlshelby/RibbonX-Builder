Namespace RibbonAttributes.Categories.Label
    Friend Class Label
        Inherits RibbonAttribute(Of String, Label)

        Public Sub New(value As String)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(Label)), GetValue())
            End Get
        End Property
    End Class
End Namespace