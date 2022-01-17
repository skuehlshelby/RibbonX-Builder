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
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Event BeforeCheckStateChange As EventHandler(Of BeforeCheckStateChangeEventArgs)

        Public Event CheckStateChanged As EventHandler(Of CheckStateChangeEventArgs)

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

        Public Property SuperTip As String Implements ISuperTip.SuperTip
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
                Dim oldValue As Boolean = Pressed

                If oldValue <> Value Then
                    Try
                        SuspendAutomaticRefreshing()

                        Dim e As BeforeCheckStateChangeEventArgs = New BeforeCheckStateChangeEventArgs(Value, oldValue)

                        RaiseEvent BeforeCheckStateChange(Me, e)

                        If Not e.Cancel Then
                            _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetPressed).SetValue(Value)

                            RaiseEvent CheckStateChanged(Me, New CheckStateChangeEventArgs(oldValue, Value))
                        End If
                    Finally
                        ResumeAutomaticRefreshing()
                    End Try
                End If
            End Set
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.ReadOnlyLookup(Of Action)(AttributeName.OnAction).GetValue().Invoke()
        End Sub

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes
        End Function

        Public Class BeforeCheckStateChangeEventArgs
            Inherits EventArgs

            Public Sub New(newValue As Boolean, oldValue As Boolean)
                Me.NewValue = newValue
                Me.OldValue = oldValue
            End Sub

            Public ReadOnly Property NewValue As Boolean

            Public ReadOnly Property OldValue As Boolean

            Public Property Cancel As Boolean

        End Class

        Public Class CheckStateChangeEventArgs
            Inherits EventArgs

            Public Sub New(newValue As Boolean, oldValue As Boolean)
                Me.NewValue = newValue
                Me.OldValue = oldValue
            End Sub

            Public ReadOnly Property NewValue As Boolean

            Public ReadOnly Property OldValue As Boolean

        End Class

    End Class

End Namespace