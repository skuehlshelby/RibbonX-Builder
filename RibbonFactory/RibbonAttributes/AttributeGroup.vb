Namespace RibbonAttributes
    Friend Class AttributeGroup
        Implements ISet(Of RibbonAttribute)
        
        Public Event AttributeChanged

        Private Sub OnValueChange()
            RaiseEvent AttributeChanged
        End Sub

        Private ReadOnly _attributes As ISet(Of RibbonAttribute) = New HashSet(Of RibbonAttribute)(RibbonAttribute.ByCategory())

        Public Function Add(item As RibbonAttribute) As Boolean Implements ISet(Of RibbonAttribute).Add
            If Not _attributes.Add(item) Then
                _attributes.Remove(item)
                _attributes.Add(item)
            End If

            Return True
        End Function

        Public Function Add(Of T)(item As RibbonAttributeReadWrite(Of T)) As Boolean
            AddHandler item.ValueChanged, AddressOf OnValueChange

            Return Add(CType(item, RibbonAttribute))
        End Function

        Public Function ReadOnlyLookup(Of T)(sampleMember As AttributeName) As RibbonAttributeReadOnly(Of T)
            Dim equatable As IEquatable(Of RibbonAttribute) = RibbonAttribute.ByCategoryMember(sampleMember)
            Return DirectCast(_attributes.First(Function(attribute) equatable.Equals(attribute)), RibbonAttributeReadOnly(Of T))
        End Function

        Public Function ReadWriteLookup(Of T)(sampleMember As AttributeName) As RibbonAttributeReadWrite(Of T)
            Dim equatable As IEquatable(Of RibbonAttribute) = RibbonAttribute.ByCategoryMember(sampleMember)
            Return DirectCast(_attributes.First(Function(attribute) equatable.Equals(attribute)), RibbonAttributeReadWrite(Of T))
        End Function

        Public Function HasAttribute(sampleMember As AttributeName) As Boolean
            Dim equatable As IEquatable(Of RibbonAttribute) = RibbonAttribute.ByCategoryMember(sampleMember)
            Return _attributes.Any(Function(attribute) equatable.Equals(attribute))
        End Function

        Public Overrides Function ToString() As String
            Return String.Join(" ", _attributes)
        End Function

        Public ReadOnly Property Count As Integer Implements ICollection(Of RibbonAttribute).Count
            Get
                Return _attributes.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of RibbonAttribute).IsReadOnly
            Get
                Return _attributes.IsReadOnly
            End Get
        End Property

        Public Sub UnionWith(other As IEnumerable(Of RibbonAttribute)) Implements ISet(Of RibbonAttribute).UnionWith
            _attributes.UnionWith(other)
        End Sub

        Public Sub IntersectWith(other As IEnumerable(Of RibbonAttribute)) Implements ISet(Of RibbonAttribute).IntersectWith
            _attributes.IntersectWith(other)
        End Sub

        Public Sub ExceptWith(other As IEnumerable(Of RibbonAttribute)) Implements ISet(Of RibbonAttribute).ExceptWith
            _attributes.ExceptWith(other)
        End Sub

        Public Sub SymmetricExceptWith(other As IEnumerable(Of RibbonAttribute)) Implements ISet(Of RibbonAttribute).SymmetricExceptWith
            _attributes.SymmetricExceptWith(other)
        End Sub

        Public Sub Clear() Implements ICollection(Of RibbonAttribute).Clear
            _attributes.Clear()
        End Sub

        Public Sub CopyTo(array() As RibbonAttribute, arrayIndex As Integer) Implements ICollection(Of RibbonAttribute).CopyTo
            _attributes.CopyTo(array, arrayIndex)
        End Sub

        Private Sub ICollection_Add(item As RibbonAttribute) Implements ICollection(Of RibbonAttribute).Add
            Add(item)
        End Sub

        Public Function IsSubsetOf(other As IEnumerable(Of RibbonAttribute)) As Boolean Implements ISet(Of RibbonAttribute).IsSubsetOf
            Return _attributes.IsSubsetOf(other)
        End Function

        Public Function IsSupersetOf(other As IEnumerable(Of RibbonAttribute)) As Boolean Implements ISet(Of RibbonAttribute).IsSupersetOf
            Return _attributes.IsSupersetOf(other)
        End Function

        Public Function IsProperSupersetOf(other As IEnumerable(Of RibbonAttribute)) As Boolean Implements ISet(Of RibbonAttribute).IsProperSupersetOf
            Return _attributes.IsProperSupersetOf(other)
        End Function

        Public Function IsProperSubsetOf(other As IEnumerable(Of RibbonAttribute)) As Boolean Implements ISet(Of RibbonAttribute).IsProperSubsetOf
            Return _attributes.IsProperSubsetOf(other)
        End Function

        Public Function Overlaps(other As IEnumerable(Of RibbonAttribute)) As Boolean Implements ISet(Of RibbonAttribute).Overlaps
            Return _attributes.Overlaps(other)
        End Function

        Public Function SetEquals(other As IEnumerable(Of RibbonAttribute)) As Boolean Implements ISet(Of RibbonAttribute).SetEquals
            Return _attributes.SetEquals(other)
        End Function

        Public Function Contains(item As RibbonAttribute) As Boolean Implements ICollection(Of RibbonAttribute).Contains
            Return _attributes.Contains(item)
        End Function

        Public Function Remove(item As RibbonAttribute) As Boolean Implements ICollection(Of RibbonAttribute).Remove
            Return _attributes.Remove(item)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonAttribute) Implements IEnumerable(Of RibbonAttribute).GetEnumerator
            Return _attributes.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(_attributes, IEnumerable).GetEnumerator()
        End Function
    End Class
End Namespace
