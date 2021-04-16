Namespace RibbonAttributes.Categories.Image
    Friend Class ImageMso
        Inherits ImageBase

        Private ReadOnly _value As Enums.ImageMSO.ImageMSO

        Public Sub New(value As Enums.ImageMSO.ImageMSO)
            _value = value
        End Sub

        Public Function GetValue() As Enums.ImageMSO.ImageMSO
            Return _value
        End Function

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(ImageMso)), GetValue())
            End Get
        End Property
    End Class
End Namespace
