Namespace RibbonAttributes.Categories.Label
    Friend Class Label
        Inherits RibbonAttribute(Of String, Label)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(Label), GetValue())
            End Get
        End Property
    End Class
End Namespace