Friend Class ControlPropertyList
    Implements IEnumerable(Of IControlProperty)

    Private ReadOnly This As HashSet(Of IControlProperty) = New HashSet(Of IControlProperty)

    Friend ReadOnly Property Count As Integer
        Get
            Return This.Count
        End Get
    End Property

    Friend Function GetValue(Of TType, TAttributeGroup As Attributes.AttributeGroup)() As TType
        Return DirectCast(This.First(Function(A) TypeOf A.Governs Is TAttributeGroup), IControlProperty(Of TType)).GetValue
    End Function

    Friend Sub SetValue(Of TType, TAttributeGroup As Attributes.AttributeGroup)(ByVal Value As TType)
        With DirectCast(This.First(Function(A) TypeOf A.Governs Is TAttributeGroup), IDynamicControlProperty(Of TType))
            If Not .GetValue().Equals(Value) Then
                .SetValue(Value)
            End If
        End With
    End Sub

    Friend Function IsReadWrite(Of TAttributeGroup As Attributes.AttributeGroup, T)() As Boolean
        Return TypeOf This.First(Function(A) TypeOf A.Governs Is TAttributeGroup) Is IDynamicControlProperty(Of T)
    End Function

    Friend Function Add(item As IControlProperty) As IControlProperty
        If Not This.Add(item) Then
            This.Remove(item)
            This.Add(item)
        End If

        Return item
    End Function

    Friend Sub Clear()
        This.Clear()
    End Sub
    Friend Function Contains(item As IControlProperty) As Boolean
        Return This.Contains(item)
    End Function

    Friend Function Remove(item As IControlProperty) As Boolean
        Return This.Remove(item)
    End Function

    Friend Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return This.GetEnumerator
    End Function

    Friend Function IEnumerable_GetEnumerator() As IEnumerator(Of IControlProperty) Implements IEnumerable(Of IControlProperty).GetEnumerator
        Return This.GetEnumerator
    End Function
End Class