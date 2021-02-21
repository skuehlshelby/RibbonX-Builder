Namespace RibbonAttributes.Categories.Image
    Friend MustInherit Class ImageBase
        Inherits RibbonAttribute

        Public Overrides Function GetHashCode() As Integer
            Return GetType(ImageBase).GetHashCode()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj.GetHashCode = GetHashCode() Then
                Return TypeOf obj Is ImageBase
            Else
                Return False
            End If
        End Function
    End Class
End Namespace