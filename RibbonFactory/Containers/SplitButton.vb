Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes


Namespace Containers
    
    Public Class SplitButton
        Inherits RibbonElement
        Implements IReadOnlyCollection(Of RibbonElement)
        Implements IVisible
        Implements IEnabled
        Implements IKeyTip
        Implements ISize
        Implements IShowLabel
        
        Private ReadOnly _attributes As AttributeGroup
        
        Friend Sub New(button As Button, menu As Menu, attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            Me.Button = button
            Me.Menu = menu
            _attributes = attributes
        End Sub
        
        Friend Sub New(button As ToggleButton, menu As Menu, attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            Me.Button = button
            Me.Menu = menu
            _attributes = attributes
        End Sub
        
        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(Environment.NewLine, $"<{NameOf(SplitButton).CamelCase()} {String.Join(" ", _attributes) }>",
                                String.Join(Environment.NewLine, WrapComponentsAsIEnumerable()), $"</{NameOf(SplitButton).CamelCase()}>")
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
                Return _attributes.Lookup(Of Size).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSize).SetValue(value)
                
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowLabel).SetValue(value)
                
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of RibbonElement).Count
            Get
                Return 2
            End Get
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