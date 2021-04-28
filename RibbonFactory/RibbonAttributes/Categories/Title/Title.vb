Namespace RibbonAttributes.Categories.Title

    Public Class Title
        Inherits RibbonAttribute(Of String, Title)

        Public Sub New(value As String)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(Title).CamelCase(), GetValue())
            End Get
        End Property
    End Class
    
End NameSpace