
Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class ComboBox
        Inherits ContainerOfItems
        Implements IComboBox

        Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

#Region "Events"

        Public Custom Event Changing As EventHandler(Of CancelableEventArgs(Of String)) Implements IComboBox.Changing
            AddHandler(value As EventHandler(Of CancelableEventArgs(Of String)))
                Attributes.AddHandler(NameOf(Changing), value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of CancelableEventArgs(Of String)))
                Attributes.RemoveHandler(NameOf(Changing), value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As CancelableEventArgs(Of String))
                Attributes.RaiseEvent(NameOf(Changing), sender, e)
            End RaiseEvent
        End Event

        Public Custom Event Changed As EventHandler(Of EventArgs(Of String)) Implements IComboBox.Changed
            AddHandler(value As EventHandler(Of EventArgs(Of String)))
                Attributes.AddHandler(NameOf(Changed), value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of EventArgs(Of String)))
                Attributes.RemoveHandler(NameOf(Changed), value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As EventArgs(Of String))
                Attributes.RaiseEvent(NameOf(Changed), sender, e)
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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Category.KeyTip)
            End Get
            Set
                Attributes.Set(Value, Category.KeyTip)
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

#Region "MaxLength"

        Public ReadOnly Property MaxLength As Byte Implements IMaxLength.MaxLength
            Get
                Return Attributes.Get(Category.MaxLength)
            End Get
        End Property

#End Region

#Region "Action"

        Public Property Text As String Implements IText.Text
            Get
                Return Attributes.Get(Category.Text)
            End Get
            Set
                Using updateBlock As IDisposable = SuspendRefreshing()
                    Dim initialValue As String = Text

                    If Not initialValue.Equals(Value, StringComparison.OrdinalIgnoreCase) AndAlso Value.Length <= MaxLength Then

                        Dim e As CancelableEventArgs(Of String) = New CancelableEventArgs(Of String)(initialValue, Value)

                        RaiseEvent Changing(Me, e)

                        If Not e.IsCancelled Then
                            Attributes.Set(e.Future, Category.Text)

                            RaiseEvent Changed(Me, New EventArgs(Of String)(e.Future))
                        End If
                    End If
                End Using
            End Set
        End Property

#End Region

        Public Overrides Function Clone() As Object
            Return New ComboBox(CType(Attributes.Clone(), IPropertyCollection), Tag)
        End Function

    End Class

End Namespace