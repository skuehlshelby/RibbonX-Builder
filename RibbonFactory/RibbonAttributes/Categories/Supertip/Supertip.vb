Namespace RibbonAttributes.Categories.Supertip
    Friend Class Supertip
        Inherits RibbonAttribute(Of String, Supertip)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(Supertip), GetValue())
            End Get
        End Property
    End Class
End Namespace