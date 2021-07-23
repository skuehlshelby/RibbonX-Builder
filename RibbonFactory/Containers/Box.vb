Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes

Namespace Containers

    Public NotInheritable Class Box
        Inherits RibbonElement
        Implements IList(Of RibbonElement)
        Implements IVisible

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _items As IList(Of RibbonElement) = New List(Of RibbonElement)

        Friend Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = buttonAttributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(Environment.NewLine, $"<{NameOf(Box).ToLower()} {String.Join(" ", _attributes) }>",
                                String.Join(Environment.NewLine, _items), $"</{NameOf(Box).ToLower()}>")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.Visible).SetValue(value)
            End Set
        End Property

        Default Public Property Item(index As Integer) As RibbonElement Implements IList(Of RibbonElement).Item
            Get
                Return _items(index)
            End Get
            Set
                _items(index) = Value
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of RibbonElement).Count
            Get
                Return _items.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of RibbonElement).IsReadOnly
            Get
                Return _items.IsReadOnly
            End Get
        End Property

        Public Sub Insert(index As Integer, element As RibbonElement) Implements IList(Of RibbonElement).Insert
            _items.Insert(index, element)
        End Sub

        Public Sub RemoveAt(index As Integer) Implements IList(Of RibbonElement).RemoveAt
            _items.RemoveAt(index)
        End Sub

        Public Sub Add(element As RibbonElement) Implements ICollection(Of RibbonElement).Add
            _items.Add(element)
        End Sub

        Public Sub Clear() Implements ICollection(Of RibbonElement).Clear
            _items.Clear()
        End Sub

        Public Sub CopyTo(array() As RibbonElement, arrayIndex As Integer) Implements ICollection(Of RibbonElement).CopyTo
            _items.CopyTo(array, arrayIndex)
        End Sub

        Public Function IndexOf(element As RibbonElement) As Integer Implements IList(Of RibbonElement).IndexOf
            Return _items.IndexOf(element)
        End Function

        Public Function Contains(element As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Contains
            Return _items.Contains(element)
        End Function

        Public Function Remove(element As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Remove
            Return _items.Remove(element)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(_items, IEnumerable).GetEnumerator()
        End Function
    End Class
End Namespace