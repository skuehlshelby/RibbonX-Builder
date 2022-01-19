Imports System.Runtime.InteropServices

Namespace RibbonAttributes
	Friend Class AttributeSet
		Implements ISet(Of IRibbonAttribute)

		Public Event AttributeChanged()

		Private Sub OnValueChange()
			RaiseEvent AttributeChanged()
		End Sub

		Private ReadOnly _attributes As ISet(Of IRibbonAttribute)

		Public Sub New()
			_attributes = New HashSet(Of IRibbonAttribute)()
		End Sub

		Public Sub New(attributes As IEnumerable(Of IRibbonAttribute))
			_attributes = New HashSet(Of IRibbonAttribute)()

			For Each attribute As IRibbonAttribute In attributes
				Add(attribute)
			Next
		End Sub

		Public Function Add(item As IRibbonAttribute) As Boolean Implements ISet(Of IRibbonAttribute).Add
			While Contains(item)
				Remove(item)
			End While

			_attributes.Add(item)

			AddHandler item.ValueChanged, AddressOf OnValueChange

			Return True
		End Function

		Public Function TryLookup(sampleMember As AttributeCategory, <Out> ByRef attribute As IRibbonAttribute) As Boolean
			attribute = _attributes.FirstOrDefault(Function(a) a.IsExclusiveWith(sampleMember))

			Return attribute IsNot Nothing
		End Function

		Public Function TryLookup(Of T)(sampleMember As AttributeCategory, <Out> ByRef attribute As IRibbonAttributeReadOnly(Of T)) As Boolean
			Dim value As IRibbonAttribute = Nothing

			If TryLookup(sampleMember, value) Then
				attribute = TryCast(value, IRibbonAttributeReadOnly(Of T))
				Return attribute IsNot Nothing
			Else
				attribute = Nothing
				Return False
			End If
		End Function

		Public Function TryLookup(Of T)(sampleMember As AttributeCategory, <Out> ByRef attribute As IRibbonAttributeReadWrite(Of T)) As Boolean
			Dim value As IRibbonAttribute = Nothing

			If TryLookup(sampleMember, value) Then
				attribute = TryCast(value, IRibbonAttributeReadWrite(Of T))
				Return attribute IsNot Nothing
			Else
				attribute = Nothing
				Return False
			End If
		End Function

		Public Function Lookup(sampleMember As AttributeCategory) As IRibbonAttribute
			Dim value As IRibbonAttribute = Nothing

			If TryLookup(sampleMember, value) Then
				Return value
			Else
				Throw New InvalidOperationException($"Couldn't find an attribute with the category '{sampleMember}'.")
			End If
		End Function

		Public Function ReadOnlyLookup(Of T)(sampleMember As AttributeName) As IRibbonAttributeReadOnly(Of T)
			Return DirectCast(_attributes.First(Function(attribute) attribute.IsExclusiveWith(sampleMember)), IRibbonAttributeReadOnly(Of T))
		End Function

		Public Function ReadOnlyLookup(Of T)(category As AttributeCategory) As IRibbonAttributeReadOnly(Of T)
			Return DirectCast(Lookup(category), IRibbonAttributeReadOnly(Of T))
		End Function

		Public Function ReadWriteLookup(Of T)(sampleMember As AttributeName) As IRibbonAttributeReadWrite(Of T)
			Return DirectCast(_attributes.First(Function(attribute) attribute.IsExclusiveWith(sampleMember)), IRibbonAttributeReadWrite(Of T))
		End Function

		Public Function ReadWriteLookup(Of T)(category As AttributeCategory) As IRibbonAttributeReadWrite(Of T)
			Dim value As IRibbonAttributeReadWrite(Of T) = TryCast(Lookup(category), IRibbonAttributeReadWrite(Of T))

			If value Is Nothing Then
				Throw New InvalidOperationException($"Cannot write to '{category}', as it is read-only.")
			Else
				Return value
			End If
		End Function

		Public Function HasAttribute(sampleMember As AttributeName) As Boolean
			Return _attributes.Any(Function(attribute) attribute.IsExclusiveWith(sampleMember))
		End Function

		Public Function HasAttribute(sampleMember As AttributeCategory) As Boolean
			Return _attributes.Any(Function(attribute) attribute.IsExclusiveWith(sampleMember))
		End Function

		Public Overrides Function ToString() As String
			Return String.Join(" ", _attributes)
		End Function

		Public ReadOnly Property Count As Integer Implements ICollection(Of IRibbonAttribute).Count
			Get
				Return _attributes.Count
			End Get
		End Property

		Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of IRibbonAttribute).IsReadOnly
			Get
				Return _attributes.IsReadOnly
			End Get
		End Property

		Public Sub UnionWith(other As IEnumerable(Of IRibbonAttribute)) Implements ISet(Of IRibbonAttribute).UnionWith
			_attributes.UnionWith(other)
		End Sub

		Public Sub IntersectWith(other As IEnumerable(Of IRibbonAttribute)) Implements ISet(Of IRibbonAttribute).IntersectWith
			_attributes.IntersectWith(other)
		End Sub

		Public Sub ExceptWith(other As IEnumerable(Of IRibbonAttribute)) Implements ISet(Of IRibbonAttribute).ExceptWith
			_attributes.ExceptWith(other)
		End Sub

		Public Sub SymmetricExceptWith(other As IEnumerable(Of IRibbonAttribute)) Implements ISet(Of IRibbonAttribute).SymmetricExceptWith
			_attributes.SymmetricExceptWith(other)
		End Sub

		Public Sub Clear() Implements ICollection(Of IRibbonAttribute).Clear
			While _attributes.Count > 0
				Remove(_attributes(0))
			End While
		End Sub

		Public Sub CopyTo(array() As IRibbonAttribute, arrayIndex As Integer) Implements ICollection(Of IRibbonAttribute).CopyTo
			_attributes.CopyTo(array, arrayIndex)
		End Sub

		Private Sub ICollection_Add(item As IRibbonAttribute) Implements ICollection(Of IRibbonAttribute).Add
			Add(item)
		End Sub

		Public Function IsSubsetOf(other As IEnumerable(Of IRibbonAttribute)) As Boolean Implements ISet(Of IRibbonAttribute).IsSubsetOf
			Return _attributes.IsSubsetOf(other)
		End Function

		Public Function IsSupersetOf(other As IEnumerable(Of IRibbonAttribute)) As Boolean Implements ISet(Of IRibbonAttribute).IsSupersetOf
			Return _attributes.IsSupersetOf(other)
		End Function

		Public Function IsProperSupersetOf(other As IEnumerable(Of IRibbonAttribute)) As Boolean Implements ISet(Of IRibbonAttribute).IsProperSupersetOf
			Return _attributes.IsProperSupersetOf(other)
		End Function

		Public Function IsProperSubsetOf(other As IEnumerable(Of IRibbonAttribute)) As Boolean Implements ISet(Of IRibbonAttribute).IsProperSubsetOf
			Return _attributes.IsProperSubsetOf(other)
		End Function

		Public Function Overlaps(other As IEnumerable(Of IRibbonAttribute)) As Boolean Implements ISet(Of IRibbonAttribute).Overlaps
			Return _attributes.Overlaps(other)
		End Function

		Public Function SetEquals(other As IEnumerable(Of IRibbonAttribute)) As Boolean Implements ISet(Of IRibbonAttribute).SetEquals
			Return _attributes.SetEquals(other)
		End Function

		Public Function Contains(item As IRibbonAttribute) As Boolean Implements ICollection(Of IRibbonAttribute).Contains
			Return _attributes.Contains(item)
		End Function

		Public Function Remove(item As IRibbonAttribute) As Boolean Implements ICollection(Of IRibbonAttribute).Remove
			If _attributes.Remove(item) Then
				RemoveHandler item.ValueChanged, AddressOf OnValueChange
				Return True
			Else
				Return False
			End If
		End Function

		Public Function Remove(category As AttributeCategory) As Boolean
			Dim value As IRibbonAttribute = Nothing

			If TryLookup(category, value) Then
				Return Remove(value)
			Else
				Return False
			End If
		End Function

		Public Function GetEnumerator() As IEnumerator(Of IRibbonAttribute) Implements IEnumerable(Of IRibbonAttribute).GetEnumerator
			Return _attributes.GetEnumerator()
		End Function

		Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
			Return DirectCast(_attributes, IEnumerable).GetEnumerator()
		End Function

	End Class

End Namespace
