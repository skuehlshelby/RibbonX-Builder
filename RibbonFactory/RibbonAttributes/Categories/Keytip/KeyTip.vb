Namespace RibbonAttributes.Categories.KeyTip
    Friend Class KeyTip
        Inherits RibbonAttribute(Of Char, KeyTip)

        Public Sub New(Value As Char)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(KeyTip), GetValue())
            End Get
        End Property
    End Class
End Namespace