Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes


Namespace Containers
    
    Public NotInheritable Class SplitButton
        Inherits RibbonElement
        Implements IEnumerable(Of RibbonElement)
        Implements IVisible
        Implements IEnabled
        Implements IKeyTip
        Implements ISize
        Implements IShowLabel
        
        Private ReadOnly _attributes As AttributeSet
        
        Friend Sub New(button As RibbonElement, menu As Menu, attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(tag)
            Require(Of ArgumentNullException)(button IsNot Nothing, $"Split buttons must be initialized with either a button or a toggle-button.")
            Require(Of ArgumentNullException)(menu IsNot Nothing, $"Split buttons must be initialized with a valid menu.")

            Me.Button = button
            Me.Menu = menu
            _attributes = attributes
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
                    String.Join(Environment.NewLine, $"<splitButton { _attributes }>",
                                String.Join(Environment.NewLine, WrapComponentsAsIEnumerable()), "</splitButton>")
            End Get
        End Property
        
        Public ReadOnly Property Button As RibbonElement
        
        Public ReadOnly Property Menu As Menu
        
        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
                
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetEnabled).SetValue(Value)
                
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)
                
            End Set
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeName.Size).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of ControlSize)(AttributeName.GetSize).SetValue(Value)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
            End Set
        End Property

        Private Function WrapComponentsAsIEnumerable() As IEnumerable(Of RibbonElement)
            Return New RibbonElement() { Button, Menu }
        End Function
        
        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return WrapComponentsAsIEnumerable().GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return WrapComponentsAsIEnumerable().GetEnumerator()
        End Function

    End Class
    
End NameSpace