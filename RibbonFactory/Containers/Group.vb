Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes

Namespace Containers

    Public Class Group
        Inherits RibbonElement
        Implements ILabel
        Implements IEnable
        Implements IVisible
        Implements ICollection(Of RibbonElement)

        Private ReadOnly _items As List(Of RibbonElement) = New List(Of RibbonElement)()

        Public Sub New(buttonAttributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(buttonAttributes, tag)
        End Sub
        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(vbNewLine + vbTab, "<group>",
                                String.Join(vbNewLine + vbTab, _items.Select(Function(g) g.XML)), "</group>")
            End Get
        End Property

        Public ReadOnly Property Count As Integer Implements ICollection(Of RibbonElement).Count
            Get
                Return _items.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of RibbonElement).IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As String)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements ILabel.ShowLabel
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Boolean)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Boolean)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnable.Enabled
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Boolean)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Sub Add(ParamArray items AS RibbonElement())
            _items.AddRange(items)
        End Sub

        Public Sub Add(item As RibbonElement) Implements ICollection(Of RibbonElement).Add
            _items.Add(item)
        End Sub

        Public Sub Clear() Implements ICollection(Of RibbonElement).Clear
            _items.Clear()
        End Sub

        Public Sub CopyTo(array() As RibbonElement, arrayIndex As Integer) Implements ICollection(Of RibbonElement).CopyTo
            _items.CopyTo(array, arrayIndex)
        End Sub

        Public Function Contains(item As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Contains
            Return _items.Contains(item)
        End Function

        Public Function Remove(item As RibbonElement) As Boolean Implements ICollection(Of RibbonElement).Remove
            Return _items.Remove(item)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Group
        End Function
    End Class
End NameSpace