
Namespace RibbonAttributes.Categories.BoxStyle

    Public Class BoxStyle
        Inherits RibbonAttribute(Of Enums.BoxStyle, BoxStyle)
    
        Public Sub New(value As Enums.BoxStyle)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(BoxStyle)), GetValue())
            End Get
        End Property
    End Class
End NameSpace