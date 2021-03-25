Namespace RibbonAttributes.Categories.Enabled
    Friend Class Enabled
        Inherits RibbonAttribute(Of Boolean, Enabled)

        Public Sub New(Value As Boolean)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(Enabled)), CamelCase(GetValue()))
            End Get
        End Property
    End Class
End Namespace