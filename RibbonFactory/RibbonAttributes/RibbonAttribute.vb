Namespace RibbonAttributes
    Public MustInherit Class RibbonAttribute(Of T, K As RibbonAttribute)
        Inherits RibbonAttribute

        Protected Value As T
        Private ReadOnly _hashCode As Integer

        Protected Sub New(value As T)
            Me.Value = value
            _hashCode = GetType(K).GetHashCode()
        End Sub

        Public Function GetValue() As T
            Return Value
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _hashCode
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is K
        End Function
    End Class
End Namespace