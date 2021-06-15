Namespace RibbonAttributes.Categories.ChildItems.Label

    Friend NotInheritable Class GetItemLabel
        Inherits RibbonAttribute(Of String, GetItemLabel)

        Public Sub New(callback As FromControlAndIndex(Of String))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetItemLabel)), Value)
            End Get
        End Property

    End Class

End NameSpace