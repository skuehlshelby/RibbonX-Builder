Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.RibbonAttributes


Namespace Controls
    
    Public NotInheritable Class Separator
        Inherits RibbonElement
        Implements IVisible
        
        Private ReadOnly _attributes As AttributeGroup
        
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
                Return $"<{NameOf(Separator).CamelCase()} { String.Join(" ", _attributes) }/>"
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