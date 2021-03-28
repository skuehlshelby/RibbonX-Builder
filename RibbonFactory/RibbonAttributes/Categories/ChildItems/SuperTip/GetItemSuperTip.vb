Namespace RibbonAttributes.Categories.ChildItems.SuperTip

    Friend NotInheritable Class GetItemSuperTip
        Inherits RibbonAttribute(Of String, GetItemSuperTip)

        Public Sub New (callback As FromControl(Of String))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetItemSuperTip)), Value)
            End Get
        End Property

    End Class

End NameSpace