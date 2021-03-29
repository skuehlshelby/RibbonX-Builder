Namespace RibbonAttributes.Categories.ChildItems.Label

    Friend NotInheritable Class ShowItemLabel
        Inherits RibbonAttribute(Of Boolean, ShowItemLabel)

        Public Sub New (value As Boolean)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(ShowItemLabel)), CamelCase(Value))
            End Get
        End Property

    End Class

End NameSpace