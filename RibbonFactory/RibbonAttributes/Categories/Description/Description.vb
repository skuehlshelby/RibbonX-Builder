Namespace RibbonAttributes.Categories.Description
    Friend Class Description
        Inherits RibbonAttribute(Of String, Description)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(Description), GetValue())
            End Get
        End Property
    End Class
End Namespace