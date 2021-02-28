Namespace RibbonAttributes.Categories.ID
    Friend Class IdQ
        Inherits Id

        Public Sub New([nameSpace] As String, id As String)
            MyBase.New(String.Join(":", [nameSpace], id))
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return If(String.IsNullOrEmpty(value), String.Empty, String.Format(XML_TEMPLATE, CamelCase(NameOf(Id)), GetValue()))
            End Get
        End Property
    End Class
End NameSpace