
Imports System.Reflection
Imports RibbonX.Utilities

Public MustInherit Class Container(Of TRibbonElement As RibbonElement)
    Inherits RibbonElement
    Implements ICollection(Of TRibbonElement)

    Protected ReadOnly Items As ICollection(Of TRibbonElement)

    Protected Sub New(items As ICollection(Of TRibbonElement), Optional tag As Object = Nothing)
        MyBase.New(tag)
        Me.Items = If(items, New List(Of TRibbonElement)())
    End Sub

    Friend Overridable Sub Flatten(results As ICollection(Of RibbonElement))
        results.Add(Me)

        Dim genericContainer As Type = GetType(Container(Of ))

        For Each item As RibbonElement In Items
            If item.GetType().IsDerivedFromGenericType(genericContainer) Then
                Dim method As MethodBase = item.GetType().GetMethod(NameOf(Flatten), BindingFlags.Instance Or BindingFlags.NonPublic)

                method.Invoke(item, {results})
            Else
                results.Add(item)
            End If
        Next
    End Sub

#Region "Add"

    Public Sub Add(item As TRibbonElement) Implements ICollection(Of TRibbonElement).Add
        BeforeAdd(item)
        Items.Add(item)
        AfterAdd(item)
    End Sub

    Protected Overridable Sub BeforeAdd(item As TRibbonElement)
    End Sub

    Protected Overridable Sub AfterAdd(item As TRibbonElement)
    End Sub

#End Region

#Region "Clear"

    Public Sub Clear() Implements ICollection(Of TRibbonElement).Clear
        If Items.Any() Then
            BeforeClear(Items)
            Items.Clear()
            AfterClear()
        End If
    End Sub

    Protected Overridable Sub BeforeClear(currentItems As ICollection(Of TRibbonElement))
    End Sub

    Protected Overridable Sub AfterClear()
    End Sub

#End Region

#Region "Remove"

    Public Function Remove(item As TRibbonElement) As Boolean Implements ICollection(Of TRibbonElement).Remove
        BeforeRemove(item)
        Dim result As Boolean = Items.Remove(item)
        AfterRemove(item)
        Return result
    End Function

    Protected Overridable Sub BeforeRemove(item As TRibbonElement)
    End Sub

    Protected Overridable Sub AfterRemove(item As TRibbonElement)
    End Sub

#End Region

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

    Public Sub CopyTo(array() As TRibbonElement, arrayIndex As Integer) Implements ICollection(Of TRibbonElement).CopyTo
        Items.CopyTo(array, arrayIndex)
    End Sub

    Public Function Contains(item As TRibbonElement) As Boolean Implements ICollection(Of TRibbonElement).Contains
        Return Items.Contains(item)
    End Function


#Region "IEnumerable Members"

    Public Overridable Function GetEnumerator() As IEnumerator(Of TRibbonElement) Implements IEnumerable(Of TRibbonElement).GetEnumerator
        Return Items.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return DirectCast(Items, IEnumerable).GetEnumerator()
    End Function

#End Region


End Class