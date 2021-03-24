Namespace RibbonAttributes
    Public Class AttributeGroup
        Inherits HashSet(Of RibbonAttribute)

        Public Shadows Function Add(Of T As RibbonAttribute)(item As T) As T
            If Not MyBase.Add(item) Then
                Remove(item)
                MyBase.Add(item)
            End If

            Return item
        End Function

        Public Function Lookup(Of T As RibbonAttribute)() As T
            Return DirectCast(Me.Single(Function(attribute) TypeOf attribute Is T), T)
        End Function

        Public Sub MergeWith(other As AttributeGroup)
            For Each attribute As RibbonAttribute In other
                Me.Add(attribute)
            Next attribute
        End Sub
    End Class
End Namespace