Namespace RibbonAttributes.Categories.SuperTip
    Friend Class Supertip
        Inherits RibbonAttribute(Of String, Supertip)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return If(String.IsNullOrEmpty(GetValue()), String.Empty, String.Format(XML_TEMPLATE, CamelCase(NameOf(Supertip)), GetValue())) 
            End Get
        End Property
    End Class
End Namespace