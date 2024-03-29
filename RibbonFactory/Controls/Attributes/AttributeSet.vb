﻿Imports RibbonX.Utilities

Namespace Controls.Attributes
    Friend Class AttributeSet
        Implements ICloneable
        Implements ICollection(Of IRibbonAttribute)

        Public Event AttributeChanged()

        Private Sub OnValueChange()
            RaiseEvent AttributeChanged()
        End Sub

        Private ReadOnly _attributes As ISet(Of IRibbonAttribute)

        Public Sub New()
            _attributes = New HashSet(Of IRibbonAttribute)(New AttributeEqualityComparer())
        End Sub

        Public Sub New(attributes As IEnumerable(Of IRibbonAttribute))
            _attributes = New HashSet(Of IRibbonAttribute)(New AttributeEqualityComparer())

            For Each attribute As IRibbonAttribute In attributes
                Add(attribute)
            Next
        End Sub

        Public Function OverwriteWithIntersectionOf(attributes As IEnumerable(Of IRibbonAttribute)) As AttributeSet
            For Each attribute As IRibbonAttribute In attributes
                If Contains(attribute) Then
                    If attribute.Category <> AttributeCategory.IdType Then
                        Add(attribute)
                    End If
                End If
            Next

            Return Me
        End Function

        Private NotInheritable Class AttributeEqualityComparer
            Implements IEqualityComparer(Of IRibbonAttribute)

            Public Overloads Function Equals(x As IRibbonAttribute, y As IRibbonAttribute) As Boolean Implements IEqualityComparer(Of IRibbonAttribute).Equals
                Return x.Category = y.Category
            End Function

            Public Overloads Function GetHashCode(obj As IRibbonAttribute) As Integer Implements IEqualityComparer(Of IRibbonAttribute).GetHashCode
                Return obj.Category.GetHashCode()
            End Function
        End Class

#Region "Add, Remove, Clear"
        Public Sub Add(item As IRibbonAttribute) Implements ICollection(Of IRibbonAttribute).Add
            Require(Of ArgumentNullException)(item IsNot Nothing)

            Remove(item)

            _attributes.Add(item)

            AddHandler item.ValueChanged, AddressOf OnValueChange
        End Sub

        Public Function Remove(item As IRibbonAttribute) As Boolean Implements ICollection(Of IRibbonAttribute).Remove
            If _attributes.Contains(item) Then
                Dim temp As IRibbonAttribute = Lookup(item.Category)

                RemoveHandler temp.ValueChanged, AddressOf OnValueChange
            End If

            Return _attributes.Remove(item)
        End Function

        Public Sub Clear() Implements ICollection(Of IRibbonAttribute).Clear
            For Each attribute As IRibbonAttribute In _attributes
                RemoveHandler attribute.ValueChanged, AddressOf OnValueChange
            Next

            _attributes.Clear()
        End Sub

#End Region

#Region "Lookup Functions"

        Private Function Lookup(category As AttributeCategory) As IRibbonAttribute
            Dim value As IRibbonAttribute = _attributes.FirstOrDefault(Function(a) a.Category = category)

            If value IsNot Nothing Then
                Return value
            Else
                Throw New InvalidOperationException($"Couldn't find an attribute with the category '{category.Name}'.")
            End If
        End Function

        Public Function Read(Of T)() As T
            Dim results As IEnumerable(Of IRibbonAttributeReadOnly(Of T)) = _attributes.OfType(Of IRibbonAttributeReadOnly(Of T))

            If results.Count() = 1 Then
                Return results.Single().GetValue()
            Else
                Throw SeekError(Of T)()
            End If
        End Function

        Public Function Read(Of T)(category As AttributeCategory) As T
            Dim results As IEnumerable(Of IRibbonAttributeReadOnly(Of T)) = _attributes.OfType(Of IRibbonAttributeReadOnly(Of T)).Where(Function(a) a.Category = category)

            If results.Count() = 1 Then
                Return results.Single().GetValue()
            Else
                Throw SeekError(Of T)(category)
            End If
        End Function

        Public Sub Write(Of T)(value As T)
            Dim results As IEnumerable(Of IRibbonAttributeReadWrite(Of T)) = _attributes.OfType(Of IRibbonAttributeReadWrite(Of T))

            If results.Count() = 1 Then
                results.Single().SetValue(value)
            Else
                Throw SeekError(Of T)()
            End If
        End Sub

        Public Sub Write(Of T)(value As T, category As AttributeCategory)
            Dim results As IEnumerable(Of IRibbonAttributeReadWrite(Of T)) = _attributes.OfType(Of IRibbonAttributeReadWrite(Of T)).Where(Function(a) a.Category = category)

            If results.Count() = 1 Then
                results.Single().SetValue(value)
            Else
                Throw SeekError(Of T)(category)
            End If
        End Sub

        Private Function SeekError(Of T)() As Exception
            Dim results As IEnumerable(Of IRibbonAttributeReadOnly(Of T)) = _attributes.OfType(Of IRibbonAttributeReadOnly(Of T))

            Select Case results.Count()
                Case 0
                    Return New Exception($"Failed to find any values of type '{GetType(T).Name}'.")
                Case 1
                    Return New Exception($"Found one value of type '{GetType(T).Name}', but it was read-only.")
                Case Else
                    Return New Exception($"Found too many values of type '{GetType(T).Name}'. Results are ambiguous.")
            End Select
        End Function

        Private Function SeekError(Of T)(category As AttributeCategory) As Exception
            Dim typeResults As IEnumerable(Of IRibbonAttributeReadOnly(Of T)) = _attributes.OfType(Of IRibbonAttributeReadOnly(Of T))
            Dim categoryResults As IEnumerable(Of IRibbonAttribute) = _attributes.Where(Function(a) a.Category = category)

            If typeResults.Count() = 0 AndAlso categoryResults.Count() = 0 Then
                Return New Exception($"Failed to find any values of type '{GetType(T).Name}'. No values belonging to the category '{category.Name}' were found either.")
            ElseIf typeResults.Count() = 0 AndAlso categoryResults.Count() = 1 Then
                Return New Exception($"Found one value belonging to category '{category.Name}', but its type was not of type '{GetType(T).Name}'.")
            ElseIf typeResults.Count() = 1 AndAlso categoryResults.Count() = 0 Then
                Return New Exception($"Found one value of type '{GetType(T).Name}', but it was not a member of category '{category.Name}'.")
            Else
                Dim combined As IEnumerable(Of IRibbonAttribute) = typeResults.Cast(Of IRibbonAttribute)().Intersect(categoryResults)

                If combined.Count() = 0 Then
                    Return New Exception($"Failed to find any values of type '{GetType(T).Name}' that also belong to category '{category.Name}'.")
                ElseIf combined.Count() = 1 Then
                    Dim result As IRibbonAttribute = combined.Single()

                    If result.Category.Equals(AttributeCategory.LabelVisibility) Then
                        Return New Exception($"Found one value of type '{GetType(T).Name}' in category '{category.Name}', but it was read-only. If this was unexpected, move Show/HideLabel() after WithLabel().")
                    ElseIf result.Category.Equals(AttributeCategory.ImageVisibility) Then
                        Return New Exception($"Found one value of type '{GetType(T).Name}' in category '{category.Name}', but it was read-only. If this was unexpected, move Show/HideImage() after WithImage().")
                    Else
                        Return New Exception($"Found one value of type '{GetType(T).Name}' in category '{category.Name}', but it was read-only.")
                    End If
                Else
                    Return New Exception($"More than one value of type '{GetType(T).Name}' belonging to category '{category.Name}' was found. This represents an internal logic error. Please report to project maintainer.")
                End If
            End If
        End Function

#End Region

        Public Overrides Function ToString() As String
            Return String.Join(" ", _attributes).NoDoubleSpace()
        End Function

        Public Overloads Function ToString(category As AttributeCategory) As String
            Return Lookup(category).ToString()
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

        Public Sub CopyTo(array() As IRibbonAttribute, arrayIndex As Integer) Implements ICollection(Of IRibbonAttribute).CopyTo
            _attributes.CopyTo(array, arrayIndex)
        End Sub

        Public Function Contains(item As IRibbonAttribute) As Boolean Implements ICollection(Of IRibbonAttribute).Contains
            Return _attributes.Contains(item)
        End Function

        Public Function Contains(category As AttributeCategory) As Boolean
            Return _attributes.Any(Function(a) a.Category = category)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of IRibbonAttribute) Implements IEnumerable(Of IRibbonAttribute).GetEnumerator
            Return _attributes.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(_attributes, IEnumerable).GetEnumerator()
        End Function

        Public Function Clone() As AttributeSet
            Return New AttributeSet(_attributes.OfType(Of ICloneable)().Select(Function(a) a.Clone()).OfType(Of IRibbonAttribute)())
        End Function

        Private Function CloneII() As Object Implements ICloneable.Clone
            Return Clone()
        End Function

    End Class

End Namespace