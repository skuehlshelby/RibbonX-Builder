Namespace RibbonAttributes.Categories.ID
    Friend Class Id
        Inherits RibbonAttribute(Of String, Id)

        Public Sub New (value As String)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return If(String.IsNullOrEmpty(value), String.Empty, String.Format(XML_TEMPLATE, CamelCase(NameOf(Id)), GetValue())) 
            End Get
        End Property
    End Class
End NameSpace