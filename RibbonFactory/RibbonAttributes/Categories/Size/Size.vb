Imports RibbonFactory.Enums

Namespace RibbonAttributes.Categories.Size
    Friend Class Size
        Inherits RibbonAttribute(Of ControlSize, Size)

        Public Sub New(Value As ControlSize)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(Size)), GetValue())
            End Get
        End Property
    End Class
End Namespace