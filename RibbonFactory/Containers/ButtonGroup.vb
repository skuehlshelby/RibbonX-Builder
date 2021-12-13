
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Containers
    
    Public NotInheritable Class ButtonGroup
        Inherits RibbonElement
        Implements IVisible
        Implements IEnumerable(Of RibbonElement)

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _items As IEnumerable(Of RibbonElement)
        
        Friend Sub New(items As IEnumerable(Of RibbonElement), attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
            _items = items
            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub
        
        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(Environment.NewLine, $"<buttonGroup { _attributes }>",
                                String.Join(Environment.NewLine, _items), "</buttonGroup")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
                
            End Set
        End Property
        
        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _items.GetEnumerator()
        End Function

    End Class
    
End NameSpace