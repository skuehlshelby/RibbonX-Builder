Imports RibbonX.Builders
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Utilities
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class CheckBox
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IDescription
        Implements IKeyTip
        Implements IChecked
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeToggleManager As EventManager(Of BeforeToggleChangeEventArgs)
        Private ReadOnly _onToggleManager As EventManager(Of ToggleChangeEventArgs)

        Public Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)

            _beforeToggleManager = New EventManager(Of BeforeToggleChangeEventArgs)($"{NameOf(CheckBox)}.{NameOf(BeforeStateChange)}")
            _onToggleManager = New EventManager(Of ToggleChangeEventArgs)($"{NameOf(CheckBox)}.{NameOf(StateChanged)}")
        End Sub

#Region "Events"

        Public Custom Event BeforeStateChange As EventHandler(Of BeforeToggleChangeEventArgs)
            AddHandler(value As EventHandler(Of BeforeToggleChangeEventArgs))
                _beforeToggleManager.Add(_attributes, value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of BeforeToggleChangeEventArgs))
                _beforeToggleManager.Remove(_attributes, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As BeforeToggleChangeEventArgs)
                _beforeToggleManager.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

        Public Custom Event StateChanged As EventHandler(Of ToggleChangeEventArgs)
            AddHandler(value As EventHandler(Of ToggleChangeEventArgs))
                _onToggleManager.Add(_attributes, value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of ToggleChangeEventArgs))
                _onToggleManager.Remove(_attributes, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As ToggleChangeEventArgs)
                _onToggleManager.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

#End Region

#Region "Display Text"

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Get(Of String)(AttributeCategory.Description)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Description)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Get(Of KeyTip)()
            End Get
            Set
                _attributes.Set(Value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Get(Of String)(AttributeCategory.Label)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Get(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Get(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.SuperTip)
            End Set
        End Property

#End Region

#Region "Enabled and Visible"

        Public Property Enabled As Boolean Implements IEnabled.Enabled

            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Enabled)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible

            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Visibility)

            End Set
        End Property

#End Region

#Region "Action"

        Public Property IsChecked As Boolean Implements IChecked.IsChecked
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Pressed)
            End Get
            Set
                Using updateBlock As IDisposable = RefreshSuspension()
                    Dim e As BeforeToggleChangeEventArgs = New BeforeToggleChangeEventArgs(IsChecked)

                    RaiseEvent BeforeStateChange(Me, e)

                    If Not e.IsCancelled Then
                        _attributes.Set(e.FutureState, AttributeCategory.Pressed)

                        RaiseEvent StateChanged(Me, New ToggleChangeEventArgs(e.FutureState))
                    End If
                End Using
            End Set
        End Property

#End Region

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

        Public Overrides Function Clone() As Object
            Return New CheckBox(Attributes.Clone(), Tag)
        End Function

    End Class

End Namespace