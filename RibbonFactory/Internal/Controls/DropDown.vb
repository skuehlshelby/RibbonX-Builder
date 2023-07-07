Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes
Imports RibbonX.Utilities

Namespace Controls

    Friend NotInheritable Class DropDown
        Inherits ContainerOfItems
        Implements IDropDown

        Private ReadOnly blank As IItem

        Public Sub New(attributes As IPropertyCollection, buttons() As IButton, blank As IItem, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
            Me.Buttons = buttons
            Me.blank = blank
        End Sub

        Public ReadOnly Property Buttons As ICollection(Of IButton) Implements IDropDown.Buttons

#Region "Events"

        Public Custom Event BeforeSelectionChange As EventHandler(Of CancelableEventArgs(Of IItem)) Implements IDropDown.SelectionChanging
            AddHandler(value As EventHandler(Of CancelableEventArgs(Of IItem)))
                Attributes.AddHandler(NameOf(BeforeSelectionChange), value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of CancelableEventArgs(Of IItem)))
                Attributes.RemoveHandler(NameOf(BeforeSelectionChange), value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As CancelableEventArgs(Of IItem))
                Attributes.RaiseEvent(NameOf(BeforeSelectionChange), sender, e)
            End RaiseEvent
        End Event

        Public Custom Event SelectionChanged As EventHandler(Of EventArgs(Of IItem)) Implements IDropDown.SelectionChanged
            AddHandler(value As EventHandler(Of EventArgs(Of IItem)))
                Attributes.AddHandler(NameOf(SelectionChanged), value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of EventArgs(Of IItem)))
                Attributes.RemoveHandler(NameOf(SelectionChanged), value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As EventArgs(Of IItem))
                Attributes.RaiseEvent(NameOf(SelectionChanged), sender, e)
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

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Get(Category.ScreenTip)
            End Get
            Set
                Attributes.Set(Value, Category.ScreenTip)
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

#Region "Action"

        Public Property Selected As IItem Implements ISelect.Selected
            Get
                Dim value As IItem = Attributes.Get(Category.SelectedItemPosition)

                If value IsNot Nothing Then
                    Return value
                Else
                    If Items.Any() Then
                        Return Items.First()
                    Else
                        Return blank
                    End If
                End If
            End Get
            Set
                Using updateBlock As IDisposable = SuspendRefreshing()
                    Dim e As CancelableEventArgs(Of IItem) = New CancelableEventArgs(Of IItem)(Selected, Value)

                    RaiseEvent BeforeSelectionChange(Me, e)

                    If Not e.IsCancelled Then
                        Attributes.Set(e.Future, Category.SelectedItemPosition)

                        RaiseEvent SelectionChanged(Me, New EventArgs(Of IItem)(e.Future))
                    End If
                End Using
            End Set
        End Property

        Public Property SelectedItemIndex As Integer Implements ISelect.SelectedItemIndex
            Get
                Return Math.Max(Items.IndexOf(Selected), 0)
            End Get
            Set
                Selected = Items(Value)
            End Set
        End Property

#End Region

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            If Buttons.Any() Then
                Return $"<dropDown {Attributes.ToXml(exclude)}>{String.Join(Environment.NewLine, Buttons.Select(Function(btn) btn.ToXml(ExcludedAttributes.Size)))}</dropDown>"
            Else
                Return $"<dropDown {Attributes.ToXml(exclude)}/>"
            End If
        End Function

        Public Overrides Function Clone() As Object
            Return New DropDown(DirectCast(Attributes.Clone(), IPropertyCollection), Buttons.Select(Function(b) b.Clone()).OfType(Of IButton)().ToArray(), DirectCast(blank.Clone(), IItem), Nothing)
        End Function

    End Class

End Namespace