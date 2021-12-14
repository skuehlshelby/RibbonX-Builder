Imports System.Reflection
Imports RibbonFactory.Utilities

Public MustInherit Class Container(Of TRibbonElement As RibbonElement)
    Inherits RibbonElement
    Implements ICollection(Of TRibbonElement)
    
    Protected ReadOnly Items As ICollection(Of TRibbonElement)

    Protected Sub New(items As ICollection(Of TRibbonElement), Optional tag As Object = Nothing)
        MyBase.New(tag)
        Me.Items = If(items, New TRibbonElement() {})
    End Sub

    Friend Overridable Sub Flatten(results As ICollection(Of RibbonElement))
        results.Add(Me)

        Dim genericContainer As Type = GetType(Container(Of ))

        For Each item As RibbonElement In Items
            If item.GetType().IsDerivedFromGenericType(genericContainer) Then
                Dim method As MethodBase = item.GetType().GetMethod("Flatten", BindingFlags.Instance Or BindingFlags.NonPublic)

                method.Invoke(item, New Object(){results})
            Else
                results.Add(item)
            End If
        Next
    End Sub

#Region "ICollection Members"

    Public ReadOnly Property Count As Integer Implements ICollection(Of TRibbonElement).Count
        Get
            Return Items.Count
        End Get
    End Property

    Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of TRibbonElement).IsReadOnly
        Get
            Return Items.IsReadOnly
        End Get
    End Property

    Public Overridable Sub Add(item As TRibbonElement) Implements ICollection(Of TRibbonElement).Add
        Items.Add(item)
        RefreshNeeded()
    End Sub

    Public Overridable Sub Clear() Implements ICollection(Of TRibbonElement).Clear
        If Items.Any() Then
            Items.Clear()
            RefreshNeeded()
        End If
    End Sub

    Public Sub CopyTo(array() As TRibbonElement, arrayIndex As Integer) Implements ICollection(Of TRibbonElement).CopyTo
        Items.CopyTo(array, arrayIndex)
    End Sub

    Public Function Contains(item As TRibbonElement) As Boolean Implements ICollection(Of TRibbonElement).Contains
        Return Items.Contains(item)
    End Function

    Public Overridable Overloads Function Remove(item As TRibbonElement) As Boolean Implements ICollection(Of TRibbonElement).Remove
        If Items.Remove(item) Then
            RefreshNeeded()
            Return True
        Else
            Return False
        End If
    End Function

    Public Overridable Overloads Function Remove(selector As Func(Of TRibbonElement, Boolean)) As Boolean
        If Items.Any(selector) Then
            Items.Remove(Items.Single(selector))
            RefreshNeeded()
            Return True
        Else
            Return False
        End If
    End Function

    Public Overridable Overloads Sub Replace(original As TRibbonElement, replacement As TRibbonElement)
        If Items.Remove(original) Then
            Items.Add(replacement)
            RefreshNeeded()
        End If
    End Sub

    Public Overridable Overloads Sub Replace(selector As Func(Of TRibbonElement, Boolean), replacement As TRibbonElement)
        If Items.Any(selector) Then
            Items.Remove(Items.Single(selector))
            Items.Add(replacement)
            RefreshNeeded()
        End If
    End Sub

    Public Overridable Function GetEnumerator() As IEnumerator(Of TRibbonElement) Implements IEnumerable(Of TRibbonElement).GetEnumerator
        Return Items.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return DirectCast(Items, IEnumerable).GetEnumerator()
    End Function

#End Region

End Class