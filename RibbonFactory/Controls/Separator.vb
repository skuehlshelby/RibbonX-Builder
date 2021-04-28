Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Visible

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
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<{NameOf(Separator).CamelCase()} { String.Join(" ", _attributes) }/>"
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
    End Class
    
End NameSpace