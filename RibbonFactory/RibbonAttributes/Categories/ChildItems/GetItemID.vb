Namespace RibbonAttributes.Categories.ChildItems
    
    Friend NotInheritable Class GetItemID
        Inherits RibbonAttribute(Of String, GetItemID)

        Public Sub New (callback As FromControl(Of String))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetItemID)), Value)
            End Get
        End Property

    End Class

End NameSpace