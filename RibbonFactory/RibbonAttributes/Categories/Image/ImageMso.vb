Namespace RibbonAttributes.Categories.Image
    Friend Class ImageMso
        Inherits ImageBase

        Private Value As Enums.ImageMSO

        Public Sub New(Value As Enums.ImageMSO)
            Me.Value = Value
        End Sub

        Public Function GetValue() As Enums.ImageMSO
            Return Value
        End Function

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(ImageMso), GetValue())
            End Get
        End Property
    End Class
End Namespace
