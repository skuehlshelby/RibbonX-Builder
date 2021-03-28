Namespace RibbonAttributes.Categories.ChildItems

    Friend NotInheritable Class GetSelectedItemIndex
        Inherits RibbonAttribute(Of String, GetSelectedItemIndex)

        Public Sub New (callback As FromControl(Of Integer))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetSelectedItemIndex)), Value)
            End Get
        End Property

    End Class

End NameSpace