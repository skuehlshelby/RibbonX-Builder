Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Containers

    Public NotInheritable Class Box
        Inherits RibbonElement
        Implements IList(Of RibbonElement)
        Implements IVisible
        
        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _items As List(Of RibbonElement) = new List(Of RibbonElement)

        Public Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = buttonAttributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
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
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(value)
            End Set
        End Property

        Default Public Property Item(index As Integer) As RibbonElement Implements IList(Of RibbonElement).Item
            Get
                Return DirectCast(_items, IList(Of RibbonElement))(index)
            End Get
            Set
                DirectCast(_items, IList(Of RibbonElement))(index) = value
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of RibbonElement).Count
            Get
                Return DirectCast(_items, ICollection(Of RibbonElement)).Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of RibbonElement).IsReadOnly
            Get
                Return DirectCast(_items, ICollection(Of RibbonElement)).IsReadOnly
            End Get
        End Property

        Public Sub Insert(index As Integer, element As RibbonElement) Implements IList(Of RibbonElement).Insert
            DirectCast(_items, IList(Of RibbonElement)).Insert(index, element)
        End Sub

        Public Sub RemoveAt(index As Integer) Implements IList(Of RibbonElement).RemoveAt
            DirectCast(_items, IList(Of RibbonElement)).RemoveAt(index)
        End Sub

        Public Sub Add(element As RibbonElement) Implements ICollection(Of RibbonElement).Add
            DirectCast(_items, ICollection(Of RibbonElement)).Add(element)
        End Sub

        Public Sub Clear() Implements ICollection(Of RibbonElement).Clear
            DirectCast(_items, ICollection(Of RibbonElement)).Clear()
        End Sub

        Public Sub CopyTo(array() As RibbonElement, arrayIndex As Integer) Implements ICollection(Of RibbonElement).CopyTo
            DirectCast(_items, ICollection(Of RibbonElement)).CopyTo(array, arrayIndex)
        End Sub

        Public Function IndexOf(element As RibbonElement) As Integer Implements IList(Of RibbonElement).IndexOf
            Return DirectCast(_items, IList(Of RibbonElement)).IndexOf(element)
        End Function

        Public Function Contains(element As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Contains
            Return DirectCast(_items, ICollection(Of RibbonElement)).Contains(element)
        End Function

        Public Function Remove(element As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Remove
            Return DirectCast(_items, ICollection(Of RibbonElement)).Remove(element)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return DirectCast(_items, IEnumerable(Of RibbonElement)).GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(_items, IEnumerable).GetEnumerator()
        End Function
    End Class
End NameSpace