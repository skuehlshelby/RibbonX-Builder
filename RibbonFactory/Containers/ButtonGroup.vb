Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Containers
    
    Public Class ButtonGroup
        Inherits RibbonElement
        Implements IVisible
        Implements IEnumerable(Of RibbonElement)
        Implements IReadOnlyCollection(Of RibbonElement)

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _items As List(Of RibbonElement) = New List(Of RibbonElement)()
        
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
                Return _
                    String.Join(Environment.NewLine, $"<{NameOf(ButtonGroup).CamelCase()} {String.Join(" ", _attributes) }>",
                                String.Join(Environment.NewLine, _items), $"</{NameOf(ButtonGroup).CamelCase()}>")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(value)
                Refresh()
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of RibbonElement).Count
            Get
                Return DirectCast(_items, IReadOnlyCollection(Of RibbonElement)).Count
            End Get
        End Property

        Public Sub Add(control As Button)
            _items.Add(control)
        End Sub
        
        Public Sub Add(control As ToggleButton)
            _items.Add(control)
        End Sub
        
        Public Sub Add(control As Menu)
            _items.Add(control)
        End Sub
        
        'TODO: Add the rest of the supported control types
        
        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _items.GetEnumerator()
        End Function
    End Class
    
End NameSpace