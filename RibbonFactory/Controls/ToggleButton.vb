Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls.Events
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public NotInheritable Class ToggleButton
        Inherits RibbonElement
        Implements IEnabled
        Implements IScreenTip
        Implements ISuperTip
        Implements IDescription
        Implements ILabel
        Implements IShowLabel
        Implements IImage
        Implements IShowImage
        Implements ISize
        Implements IPressed
        Implements IVisible
        Implements IKeyTip
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeToggleManager As EventManager(Of BeforeToggleChangeEventArgs)
        Private ReadOnly _onToggleManager As EventManager(Of ToggleChangeEventArgs)

        Public Sub New()
            Me.New(Nothing, Nothing, Nothing)
        End Sub

        Public Sub New(configuration As Action(Of IToggleButtonBuilder), Optional tag As Object = Nothing)
            Me.New(Nothing, configuration, tag)
        End Sub

        Public Sub New(template As RibbonElement, configuration As Action(Of IToggleButtonBuilder), Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As ToggleButtonBuilder = New ToggleButtonBuilder(template)

            If configuration IsNot Nothing Then
                configuration.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            _beforeToggleManager = New EventManager(Of BeforeToggleChangeEventArgs)($"{NameOf(ToggleButton)}.{NameOf(BeforeStateChange)}")
            _onToggleManager = New EventManager(Of ToggleChangeEventArgs)($"{NameOf(ToggleButton)}.{NameOf(StateChanged)}")
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

#Region "Base Class Overrides"

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Read(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<toggleButton { _attributes }/>"
            End Get
        End Property

#End Region

#Region "Enabled and Visible"

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Enabled)

            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Visibility)

            End Set
        End Property

#End Region

#Region "DisplayText"

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Description)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Description)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Label)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.LabelVisibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.LabelVisibility)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.SuperTip)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Read(Of KeyTip)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

#End Region

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Read(Of ControlSize)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

#End Region

#Region "Action"

        Public Property Checked As Boolean Implements IPressed.Checked
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Pressed)
            End Get
            Set
                Using updateBlock As IDisposable = RefreshSuspension()
                    Dim e As BeforeToggleChangeEventArgs = New BeforeToggleChangeEventArgs(Checked)

                    RaiseEvent BeforeStateChange(Me, e)

                    If Not e.IsCancelled Then
                        _attributes.Write(Value, AttributeCategory.Pressed)

                        RaiseEvent StateChanged(Me, New ToggleChangeEventArgs(Value))
                    End If
                End Using
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Read(Of RibbonImage)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.ImageVisibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.ImageVisibility)
            End Set
        End Property

#End Region

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace