Namespace Enums.ImageMSO
    Public MustInherit Class ImageMSO
        Protected Sub New(name As String)
            Me.Name = name
        End Sub

        Public ReadOnly Property Name As String
        
        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class
End NameSpace