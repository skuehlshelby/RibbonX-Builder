Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Containers

    Public Class Box
        Inherits RibbonElement
        Implements IReadOnlyCollection(Of RibbonElement)
        Implements IVisible

        Private ReadOnly _orientation AS BoxStyle
        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _items As List(Of RibbonElement)

        Public Sub New(buttonAttributes As AttributeGroup, orientation As BoxStyle, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = buttonAttributes
            _orientation = orientation
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<box boxStyle=""{_orientation}"" {String.Join(" ", _attributes)}/>"
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

        Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of RibbonElement).Count
            Get
                Return _items.Count
            End Get
        End Property

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Box
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _items.GetEnumerator()
        End Function
    End Class
End NameSpace