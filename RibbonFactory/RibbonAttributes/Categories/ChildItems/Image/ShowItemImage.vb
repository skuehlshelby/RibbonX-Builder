Namespace RibbonAttributes.Categories.ChildItems.Image

    Friend NotInheritable Class ShowItemImage
        Inherits RibbonAttribute(Of Boolean, ShowItemImage)

        Public Sub New (value As Boolean)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(ShowItemImage)), Value)
            End Get
        End Property

    End Class

End NameSpace