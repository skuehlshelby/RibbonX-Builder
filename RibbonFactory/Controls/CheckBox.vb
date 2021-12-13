
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes


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

        Private ReadOnly _attributes As AttributeSet

        Friend Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(tag)
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
                Return $"<checkBox { _attributes }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled

            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetEnabled).SetValue(Value)
                
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible

            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
                
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
                
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Screentip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetScreentip).SetValue(Value)
            End Set
        End Property

        Public Property SuperTip As String Implements ISupertip.Supertip
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Supertip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetSupertip).SetValue(Value)
            End Set
        End Property

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Description).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetDescription).SetValue(Value)
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

        Public Property Pressed As Boolean Implements IPressed.Pressed
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.GetPressed).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetPressed).SetValue(Value)
            End Set
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.ReadOnlyLookup(Of Action)(AttributeName.OnAction).GetValue().Invoke()
        End Sub

    End Class

End Namespace