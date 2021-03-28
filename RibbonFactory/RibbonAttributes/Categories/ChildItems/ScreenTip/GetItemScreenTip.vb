Namespace RibbonAttributes.Categories.ChildItems.ScreenTip

    Friend NotInheritable Class GetItemScreenTip
        Inherits RibbonAttribute(Of String, GetItemScreenTip)

        Public Sub New (callback As FromControl(Of String))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetItemScreenTip)), Value)
            End Get
        End Property

    End Class

End NameSpace