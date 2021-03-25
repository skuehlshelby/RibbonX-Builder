Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Pressed
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Controls

    Public Class CheckBox
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IDescription
        Implements IKeyTip
        Implements IOnAction
        Implements IPressed

        Private ReadOnly _attributes As AttributeGroup

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
                Return $"<checkBox { String.Join(" ", _attributes) }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled

            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible

            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Lookup (Of SuperTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetSuperTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Lookup (Of Description).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetDescription).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Lookup (Of Categories.KeyTip.KeyTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetKeyTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Pressed As Boolean Implements IPressed.Pressed
            Get
                Return _attributes.Lookup (Of GetPressed).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetPressed).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.Lookup(Of Categories.OnAction.OnAction).GetValue()
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is CheckBox
        End Function
    End Class

End Namespace