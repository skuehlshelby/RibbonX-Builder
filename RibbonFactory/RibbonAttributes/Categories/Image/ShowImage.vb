Namespace RibbonAttributes.Categories.Image
    Friend Class ShowImage
        Inherits RibbonAttribute(Of Boolean, ShowImage)
        Public Sub New(value As Boolean)
            MyBase.New(value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Throw New NotImplementedException()
            End Get
        End Property
    End Class
End NameSpace