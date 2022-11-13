Imports RibbonX.Api
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Controls.Properties
Imports RibbonX.Controls.Utilities
Imports RibbonX.Images
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class ToggleButton
        Inherits RibbonElement
        Implements IToggleButton
        Implements IAttributeSource

        Private ReadOnly _beforeToggleManager As EventManager(Of BeforeToggleChangeEventArgs)
        Private ReadOnly _onToggleManager As EventManager(Of ToggleChangeEventArgs)

        Public Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)

            _beforeToggleManager = New EventManager(Of BeforeToggleChangeEventArgs)($"{NameOf(ToggleButton)}.{NameOf(BeforeStateChange)}")
            _onToggleManager = New EventManager(Of ToggleChangeEventArgs)($"{NameOf(ToggleButton)}.{NameOf(StateChanged)}")
        End Sub

#Region "Events"

        Public Custom Event BeforeStateChange As EventHandler(Of BeforeToggleChangeEventArgs)
            AddHandler(value As EventHandler(Of BeforeToggleChangeEventArgs))
                _beforeToggleManager.Add(Attributes, value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of BeforeToggleChangeEventArgs))
                _beforeToggleManager.Remove(Attributes, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As BeforeToggleChangeEventArgs)
                _beforeToggleManager.FireEvent(Attributes, sender, e)
            End RaiseEvent
        End Event

        Public Custom Event StateChanged As EventHandler(Of ToggleChangeEventArgs)
            AddHandler(value As EventHandler(Of ToggleChangeEventArgs))
                _onToggleManager.Add(Attributes, value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of ToggleChangeEventArgs))
                _onToggleManager.Remove(Attributes, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As ToggleChangeEventArgs)
                _onToggleManager.FireEvent(Attributes, sender, e)
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

#Region "DisplayText"

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
                Return Attributes.Get(Of KeyTip)()
            End Get
            Set
                Attributes.Set(Value)
            End Set
        End Property

#End Region

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return Attributes.Get(Of ControlSize)()
            End Get
            Set
                Attributes.Set(Value)
            End Set
        End Property

#End Region

#Region "Action"

        Public Property IsChecked As Boolean Implements IChecked.IsChecked
            Get
                Return Attributes.Get(Of Boolean)(AttributeCategory.Pressed)
            End Get
            Set
                Using updateBlock As IDisposable = RefreshSuspension()
                    Dim e As BeforeToggleChangeEventArgs = New BeforeToggleChangeEventArgs(IsChecked)

                    RaiseEvent BeforeStateChange(Me, e)

                    If Not e.IsCancelled Then
                        Attributes.Set(Value, AttributeCategory.Pressed)

                        RaiseEvent StateChanged(Me, New ToggleChangeEventArgs(Value))
                    End If
                End Using
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return Attributes.Get(Of RibbonImage)()
            End Get
            Set
                Attributes.Set(Value)
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

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return Attributes.Clone()
        End Function

        Public Overrides Function Clone() As Object
            Return New ToggleButton(Attributes.Clone(), Tag)
        End Function

    End Class

End Namespace