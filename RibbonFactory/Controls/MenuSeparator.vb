Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public NotInheritable Class MenuSeparator
        Inherits RibbonElement
        Implements ITitle
        
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
                Return $"<menuSeparator { _attributes }/>"
            End Get
        End Property

        Public Property Title As String Implements ITitle.Title
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Title).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetTitle).SetValue(value)
            End Set
        End Property

    End Class

End NameSpace