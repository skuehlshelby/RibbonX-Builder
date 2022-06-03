Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Builders
Imports RibbonX.ControlInterfaces
Imports RibbonX.Controls.Events
Imports RibbonX.Enums
Imports RibbonX.RibbonAttributes

Namespace Controls

    Public NotInheritable Class Button
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IShowLabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IKeyTip
        Implements IDescription
        Implements IImage
        Implements IClickable
        Implements IShowImage
        Implements ISize
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeClickManager As EventManager(Of CancelableEventArgs)
        Private ReadOnly _onClickManager As EventManager

        Public Sub New(Optional steps As Action(Of IButtonBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As ButtonBuilder = New ButtonBuilder(template)

            If steps IsNot Nothing Then
                steps.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            _beforeClickManager = New EventManager(Of CancelableEventArgs)($"{NameOf(Button)}.{NameOf(BeforeClick)}")
            _onClickManager = New EventManager($"{NameOf(Button)}.{NameOf(OnClick)}")
        End Sub

#Region "Events"

        Public Custom Event BeforeClick As EventHandler(Of CancelableEventArgs)
            AddHandler(value As EventHandler(Of CancelableEventArgs))
                _beforeClickManager.Add(_attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of CancelableEventArgs))
                _beforeClickManager.Remove(_attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As CancelableEventArgs)
                _beforeClickManager.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

        Public Custom Event OnClick As EventHandler
            AddHandler(value As EventHandler)
                _onClickManager.Add(_attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler)
                _onClickManager.Remove(_attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As EventArgs)
                _onClickManager.FireEvent(_attributes, sender, e)
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
                Return $"<button { _attributes }/>"
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

#Region "Display Text"

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
                Return _attributes.Read(Of KeyTip)(AttributeCategory.KeyTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.KeyTip)
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Read(Of RibbonImage)()
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Image)
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

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Read(Of ControlSize)(AttributeCategory.Size)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Size)
            End Set
        End Property

#End Region

#Region "Action"

        Public Sub Click() Implements IClickable.Click
            Dim e As CancelableEventArgs = New CancelableEventArgs()

            RaiseEvent BeforeClick(Me, e)

            If Not e.IsCancelled Then
                RaiseEvent OnClick(Me, EventArgs.Empty)
            End If
        End Sub

#End Region

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace