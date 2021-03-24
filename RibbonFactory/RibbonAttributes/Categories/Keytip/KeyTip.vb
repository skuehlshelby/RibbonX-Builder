Namespace RibbonAttributes.Categories.KeyTip
    Friend Class KeyTip
        Inherits RibbonAttribute(Of RibbonFactory.KeyTip, KeyTip)

        Public Sub New(value As RibbonFactory.KeyTip)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(KeyTip), GetValue())
            End Get
        End Property
    End Class
End Namespace