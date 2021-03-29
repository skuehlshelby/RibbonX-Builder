Namespace RibbonAttributes.Categories.ChildItems

    Friend NotInheritable Class InvalidateContentOnDrop
        Inherits RibbonAttribute(Of Boolean, InvalidateContentOnDrop)

        Public Sub New (value As Boolean)
            MyBase.New(value:= value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(InvalidateContentOnDrop)), CamelCase(Value))
            End Get
        End Property
    End Class

End NameSpace