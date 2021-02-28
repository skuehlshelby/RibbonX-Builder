Namespace RibbonAttributes.Categories.Description
    Friend Class Description
        Inherits RibbonAttribute(Of String, Description)

        Public Sub New(value As String)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return If(String.IsNullOrEmpty(Value), String.Empty, String.Format(XML_TEMPLATE, CamelCase(NameOf(Description)), GetValue()))
            End Get
        End Property
    End Class
End Namespace