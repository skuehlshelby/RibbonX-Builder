
Imports RibbonX.Controls.Base
Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class ToggleButton
        Inherits RibbonElement
        Implements IToggleButton

        Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

#Region "Events"

        Public Custom Event Checking As EventHandler(Of CancelableEventArgs(Of Boolean)) Implements IToggleButton.Changing
            AddHandler(value As EventHandler(Of CancelableEventArgs(Of Boolean)))
                Attributes.AddHandler(NameOf(Checking), value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of CancelableEventArgs(Of Boolean)))
                Attributes.RemoveHandler(NameOf(Checking), value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As CancelableEventArgs(Of Boolean))
                Attributes.RaiseEvent(NameOf(Checking), sender, e)
            End RaiseEvent
        End Event

        Public Custom Event Checked As EventHandler(Of EventArgs(Of Boolean)) Implements IToggleButton.Changed
            AddHandler(value As EventHandler(Of EventArgs(Of Boolean)))
                Attributes.AddHandler(NameOf(Checked), value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of EventArgs(Of Boolean)))
                Attributes.RemoveHandler(NameOf(Checked), value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As EventArgs(Of Boolean))
                Attributes.RaiseEvent(NameOf(Checked), sender, e)
            End RaiseEvent
        End Event

#End Region

#Region "Enabled and Visible"

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return Attributes.Get(Category.Enabled)
            End Get
            Set
                Attributes.Set(Value, Category.Enabled)

            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Get(Category.Visibility)
            End Get
            Set
                Attributes.Set(Value, Category.Visibility)

            End Set
        End Property

#End Region

#Region "DisplayText"

        Public Property Description As String Implements IDescription.Description
            Get
                Return Attributes.Get(Category.Description)
            End Get
            Set
                Attributes.Set(Value, Category.Description)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Get(Category.Label)
            End Get
            Set
                Attributes.Set(Value, Category.Label)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return Attributes.Get(Category.LabelVisibility)
            End Get
            Set
                Attributes.Set(Value, Category.LabelVisibility)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Get(Category.ScreenTip)
            End Get
            Set
                Attributes.Set(Value, Category.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return Attributes.Get(Category.SuperTip)
            End Get
            Set
                Attributes.Set(Value, Category.SuperTip)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Category.KeyTip)
            End Get
            Set
                Attributes.Set(Value, Category.KeyTip)
            End Set
        End Property

#End Region

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return Attributes.Get(Category.Size)
            End Get
            Set
                Attributes.Set(Value, Category.Size)
            End Set
        End Property

#End Region

#Region "Action"

        Public Property IsChecked As Boolean Implements IChecked.IsChecked
            Get
                Return Attributes.Get(Category.Pressed)
            End Get
            Set
                Using updateBlock As IDisposable = SuspendRefreshing()
                    Dim e As CancelableEventArgs(Of Boolean) = New CancelableEventArgs(Of Boolean)(IsChecked, Value)

                    RaiseEvent Checking(Me, e)

                    If Not e.IsCancelled Then
                        Attributes.Set(e.Future, Category.Pressed)

                        RaiseEvent Checked(Me, New EventArgs(Of Boolean)(e.Future))
                    End If
                End Using
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return Attributes.Get(Category.Image)
            End Get
            Set
                Attributes.Set(Value, Category.Image)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return Attributes.Get(Category.ImageVisibility)
            End Get
            Set
                Attributes.Set(Value, Category.ImageVisibility)
            End Set
        End Property

#End Region

        Public Overrides Function Clone() As Object
            Return New ToggleButton(CType(Attributes.Clone(), IPropertyCollection), Tag)
        End Function

    End Class

End Namespace