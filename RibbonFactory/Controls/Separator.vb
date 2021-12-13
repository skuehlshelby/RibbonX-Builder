
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Controls
    
    Public NotInheritable Class Separator
        Inherits RibbonElement
        Implements IVisible
        
        Private ReadOnly _attributes As AttributeSet
        
        Friend Sub New(buttonAttributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = buttonAttributes
            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub
        
        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<separator { _attributes }/>"
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
    End Class
    
End NameSpace