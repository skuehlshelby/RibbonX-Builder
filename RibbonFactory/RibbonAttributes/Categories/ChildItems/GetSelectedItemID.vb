Namespace RibbonAttributes.Categories.ChildItems
    
    Friend NotInheritable Class GetSelectedItemID
        Inherits RibbonAttribute(Of String, GetSelectedItemID)

        Public Sub New (callback As FromControl(Of String))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetSelectedItemID)), Value)
            End Get
        End Property

    End Class

End NameSpace