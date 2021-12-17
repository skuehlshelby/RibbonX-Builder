Namespace RibbonAttributes

    Friend Class RibbonAttributeDefault(Of T)
        Inherits RibbonAttributeReadOnly(Of T)

        Public Sub New(value As T, name As AttributeName, category As AttributeCategory)
            MyBase.New(value, name, category)
        End Sub

        Public Overrides ReadOnly Property Xml As String
            Get
                Return String.Empty
            End Get
        End Property
    End Class

End NameSpace