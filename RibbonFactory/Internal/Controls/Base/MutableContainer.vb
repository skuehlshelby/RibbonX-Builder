Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi

Friend MustInherit Class MutableContainer(Of T As IRibbonElement)
    Inherits RibbonElement
    Implements ICollection(Of T)

    Protected ReadOnly Items As ICollection(Of T) = New LinkedList(Of T)

    Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
        MyBase.New(attributes, tag)
    End Sub

    Public ReadOnly Property Count As Integer Implements ICollection(Of T).Count
        Get
            Return Items.Count
        End Get
    End Property

    Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of T).IsReadOnly
        Get
            Return Items.IsReadOnly
        End Get
    End Property

    Public Sub Add(item As T) Implements ICollection(Of T).Add
        Items.Add(item)
        AddHandler item.PropertyChanged, AddressOf OnPropertyChanged
        AddHandler item.RefreshNeeded, AddressOf OnRefreshNeeded
        TriggerRefresh()
    End Sub

    Public Sub Clear() Implements ICollection(Of T).Clear
        If Items.Any() Then
            For Each item As T In Items
                RemoveHandler item.PropertyChanged, AddressOf OnPropertyChanged
                RemoveHandler item.RefreshNeeded, AddressOf OnRefreshNeeded
            Next
            Items.Clear()
            TriggerRefresh()
        End If
    End Sub

    Public Sub CopyTo(array() As T, arrayIndex As Integer) Implements ICollection(Of T).CopyTo
        Items.CopyTo(array, arrayIndex)
    End Sub

    Public Function Contains(item As T) As Boolean Implements ICollection(Of T).Contains
        Return Items.Contains(item)
    End Function

    Public Function Remove(item As T) As Boolean Implements ICollection(Of T).Remove
        If Items.Remove(item) Then
            RemoveHandler item.PropertyChanged, AddressOf OnPropertyChanged
            RemoveHandler item.RefreshNeeded, AddressOf OnRefreshNeeded
            TriggerRefresh()
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
        Return Items.GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return DirectCast(Items, IEnumerable).GetEnumerator()
    End Function
End Class
