Namespace RibbonAttributes.Categories.ID
    Friend Class IdQ
        Inherits Id

        Public Sub New([nameSpace] As String, id As String)
            MyBase.New(String.Join(":", [nameSpace], id))
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(IdQ), GetValue())
            End Get
        End Property
    End Class
End NameSpace