Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Containers

    Public NotInheritable Class Tab
        Inherits RibbonElement
        Implements IList(Of RibbonElement)
        Implements IVisible
        Implements IKeyTip
        Implements ILabel

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _childElements As List(Of RibbonElement) = New List(Of RibbonElement)()

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Join(Environment.NewLine,$"<tab {String.Join(" ", _attributes) }>", String.Join(Environment.NewLine, _childElements), "</tab>")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup (Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetVisible).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Lookup (Of Categories.KeyTip.KeyTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetKeyTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup (Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetLabel).SetValue(value)
                Refresh()
            End Set
        End Property

        Default Public Property Item(index As Integer) As RibbonElement Implements IList(Of RibbonElement).Item
            Get
                Return DirectCast(_childElements, IList(Of RibbonElement))(index)
            End Get
            Set(value As RibbonElement)
                DirectCast(_childElements, IList(Of RibbonElement))(index) = value
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of RibbonElement).Count
            Get
                Return DirectCast(_childElements, ICollection(Of RibbonElement)).Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of RibbonElement).IsReadOnly
            Get
                Return DirectCast(_childElements, ICollection(Of RibbonElement)).IsReadOnly
            End Get
        End Property

        Public Sub Insert(index As Integer, element As RibbonElement) Implements IList(Of RibbonElement).Insert
            DirectCast(_childElements, IList(Of RibbonElement)).Insert(index, element)
        End Sub

        Public Sub RemoveAt(index As Integer) Implements IList(Of RibbonElement).RemoveAt
            DirectCast(_childElements, IList(Of RibbonElement)).RemoveAt(index)
        End Sub

        Public Sub Add(element As RibbonElement) Implements ICollection(Of RibbonElement).Add
            DirectCast(_childElements, ICollection(Of RibbonElement)).Add(element)
        End Sub

        Public Sub Clear() Implements ICollection(Of RibbonElement).Clear
            DirectCast(_childElements, ICollection(Of RibbonElement)).Clear()
        End Sub

        Public Sub CopyTo(array() As RibbonElement, arrayIndex As Integer) Implements ICollection(Of RibbonElement).CopyTo
            DirectCast(_childElements, ICollection(Of RibbonElement)).CopyTo(array, arrayIndex)
        End Sub

        Public Function AddGroup(group As Group) As Group
            _childElements.Add(group)
            Return group
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Tab
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _childElements.GetEnumerator()
        End Function

        Public Function IndexOf(element As RibbonElement) As Integer Implements IList(Of RibbonElement).IndexOf
            Return DirectCast(_childElements, IList(Of RibbonElement)).IndexOf(element)
        End Function

        Public Function Contains(element As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Contains
            Return DirectCast(_childElements, ICollection(Of RibbonElement)).Contains(element)
        End Function

        Public Function Remove(element As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Remove
            Return DirectCast(_childElements, ICollection(Of RibbonElement)).Remove(element)
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _childElements.GetEnumerator()
        End Function
    End Class
End NameSpace