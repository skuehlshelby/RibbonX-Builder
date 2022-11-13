Imports RibbonX.Builders
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Utilities
Imports RibbonX.SimpleTypes
Imports RibbonX.Images
Imports RibbonX.Api

Namespace Controls

    Friend NotInheritable Class Button
        Inherits RibbonElement
        Implements IAttributeSource
        Implements IButton

        Private ReadOnly _beforeClickManager As EventManager(Of CancelableEventArgs)
        Private ReadOnly _onClickManager As EventManager

        Public Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)

            _beforeClickManager = New EventManager(Of CancelableEventArgs)($"{NameOf(Button)}.{NameOf(BeforeClick)}")
            _onClickManager = New EventManager($"{NameOf(Button)}.{NameOf(OnClick)}")
        End Sub

#Region "Events"

        Public Custom Event BeforeClick As EventHandler(Of CancelableEventArgs)
            AddHandler(value As EventHandler(Of CancelableEventArgs))
                _beforeClickManager.Add(Attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of CancelableEventArgs))
                _beforeClickManager.Remove(Attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As CancelableEventArgs)
                _beforeClickManager.FireEvent(Attributes, sender, e)
            End RaiseEvent
        End Event

        Public Custom Event OnClick As EventHandler
            AddHandler(value As EventHandler)
                _onClickManager.Add(Attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler)
                _onClickManager.Remove(Attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As System.EventArgs)
                _onClickManager.FireEvent(Attributes, sender, e)
            End RaiseEvent
        End Event

#End Region

#Region "Enabled and Visible"

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return Attributes.Get(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Enabled)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Get(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Visibility)
            End Set
        End Property

#End Region

#Region "Display Text"

        Public Property Description As String Implements IDescription.Description
            Get
                Return Attributes.Get(Of String)(AttributeCategory.Description)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Description)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Get(Of String)(AttributeCategory.Label)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return Attributes.Get(Of Boolean)(AttributeCategory.LabelVisibility)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.LabelVisibility)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Get(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.ScreenTip)
            End Set
        End Property
        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return Attributes.Get(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.SuperTip)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Of KeyTip)(AttributeCategory.KeyTip)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.KeyTip)
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return Attributes.Get(Of RibbonImage)()
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Image)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return Attributes.Get(Of Boolean)(AttributeCategory.ImageVisibility)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.ImageVisibility)
            End Set
        End Property

#End Region

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return Attributes.Get(Of ControlSize)(AttributeCategory.Size)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Size)
            End Set
        End Property

#End Region

#Region "Action"

        Public Sub Click() Implements IClickable.Click
            Dim e As CancelableEventArgs = New CancelableEventArgs()

            RaiseEvent BeforeClick(Me, e)

            If Not e.IsCancelled Then
                RaiseEvent OnClick(Me, System.EventArgs.Empty)
            End If
        End Sub

#End Region

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return Attributes.Clone()
        End Function

        Public Overrides Function Clone() As Object
            Return New Button(Attributes.Clone(), Tag)
        End Function

    End Class

End Namespace