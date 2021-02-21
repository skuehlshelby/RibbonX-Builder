Namespace RibbonAttributes.Categories.Screentip
    Friend Class Screentip
        Inherits RibbonAttribute(Of String, Screentip)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(Screentip), GetValue())
            End Get
        End Property
    End Class
End Namespace