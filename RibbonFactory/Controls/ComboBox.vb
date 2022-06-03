Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Builders
Imports RibbonX.ControlInterfaces
Imports RibbonX.RibbonAttributes
Imports RibbonX.Controls.Events

Namespace Controls

    Public NotInheritable Class ComboBox
        Inherits Container(Of Item)
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IShowLabel
        Implements IKeyTip
        Implements IImage
        Implements IShowImage
        Implements IMaxLength
        Implements IText
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeTextChangedHelper As EventManager(Of BeforeTextChangeEventArgs)
        Private ReadOnly _textChangedHelper As EventManager(Of TextChangeEventArgs)

        Public Sub New()
            Me.New(Nothing, Nothing, Nothing)
        End Sub

        Public Sub New(configuration As Action(Of IComboBoxBuilder), Optional tag As Object = Nothing)
            Me.New(configuration, Nothing, tag)
        End Sub

        Public Sub New(configuration As Action(Of IComboBoxBuilder), template As RibbonElement, Optional tag As Object = Nothing)
            MyBase.New(New List(Of Item)(), tag)

            Dim builder As ComboBoxBuilder = New ComboBoxBuilder(template)

            If configuration IsNot Nothing Then
                configuration.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            _beforeTextChangedHelper = New EventManager(Of BeforeTextChangeEventArgs)($"{NameOf(ComboBox)}.{NameOf(BeforeTextChange)}")
            _textChangedHelper = New EventManager(Of TextChangeEventArgs)($"{NameOf(ComboBox)}.{NameOf(TextChanged)}")
        End Sub

#Region "Events"

        Public Custom Event BeforeTextChange As EventHandler(Of BeforeTextChangeEventArgs)
            AddHandler(value As EventHandler(Of BeforeTextChangeEventArgs))
                _beforeTextChangedHelper.Add(_attributes, value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of BeforeTextChangeEventArgs))
                _beforeTextChangedHelper.Remove(_attributes, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As BeforeTextChangeEventArgs)
                _beforeTextChangedHelper.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

        Public Custom Event TextChanged As EventHandler(Of TextChangeEventArgs)
            AddHandler(value As EventHandler(Of TextChangeEventArgs))
                _textChangedHelper.Add(_attributes, value)
            End AddHandler
            RemoveHandler(value As EventHandler(Of TextChangeEventArgs))
                _textChangedHelper.Remove(_attributes, value)
            End RemoveHandler
            RaiseEvent(sender As Object, e As TextChangeEventArgs)
                _textChangedHelper.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

#End Region

#Region "RibbonElement Overrides"

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Read(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<comboBox { _attributes }/>"
            End Get
        End Property

#End Region

#Region "Container Overrides"

        Protected Overrides Sub AfterAdd(item As Item)
            AddHandler item.ValueChanged, AddressOf OnChildItemChange
            RefreshNeeded()
        End Sub

        Public Overloads Sub Add(itemLabel As String, itemSupertip As String)
            Add(New Item(Sub(b) b.WithLabel(itemLabel).WithSuperTip(itemSupertip)))
        End Sub

        Protected Overrides Sub BeforeClear(currentItems As ICollection(Of Item))
            For Each item As Item In currentItems
                RemoveHandler item.ValueChanged, AddressOf OnChildItemChange
            Next
        End Sub

        Protected Overrides Sub AfterClear()
            RefreshNeeded()
        End Sub

        Protected Overrides Sub AfterRemove(item As Item)
            RemoveHandler item.ValueChanged, AddressOf OnChildItemChange
            RefreshNeeded()
        End Sub

        Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
            results.Add(Me)
        End Sub

#End Region

#Region "Container Helpers"

        Private Sub OnChildItemChange(sender As Object, e As ValueChangedEventArgs)
            RefreshNeeded()
        End Sub

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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Read(Of KeyTip)(AttributeCategory.KeyTip)
            End Get
            Set
                _attributes.Write(Value)
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

#Region "MaxLength"

        Public ReadOnly Property MaxLength As Byte Implements IMaxLength.MaxLength
            Get
                Return _attributes.Read(Of Byte)(AttributeCategory.MaxLength)
            End Get
        End Property

#End Region

#Region "Action"

        Public Property Text As String Implements IText.Text
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Text)
            End Get
            Set
                Using updateBlock As IDisposable = RefreshSuspension()
                    Dim initialValue As String = Text

                    If Not initialValue.Equals(Value, StringComparison.OrdinalIgnoreCase) AndAlso Value.Length <= MaxLength Then

                        Dim e As BeforeTextChangeEventArgs = New BeforeTextChangeEventArgs(initialValue, Value)

                        RaiseEvent BeforeTextChange(Me, e)

                        If Not e.IsCancelled Then
                            Dim updatedValue As String = e.NewText

                            _attributes.Write(updatedValue, AttributeCategory.Text)

                            RaiseEvent TextChanged(Me, New TextChangeEventArgs(updatedValue))
                        End If
                    End If
                End Using
            End Set
        End Property

#End Region

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace