Namespace RibbonAttributes.Categories.ChildItems
    Public Class ItemSize
        Inherits RibbonAttribute(Of Enums.ControlSize, ItemSize)
        
        Public Sub New(value As Enums.ControlSize)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(ItemSize)), GetValue())
            End Get
        End Property
    End Class
End NameSpace