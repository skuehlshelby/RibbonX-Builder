Imports System.Text.RegularExpressions
Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Controls.Properties
Imports RibbonX.Controls.Utilities
Imports RibbonX.Images
Imports RibbonX.SimpleTypes
Imports RibbonX.Utilities

Namespace Controls

    Public NotInheritable Class Gallery
        Inherits Container(Of Item)
        Implements IDescription
        Implements IEnabled
        Implements IVisible
        Implements IImage
        Implements IShowImage
        Implements IItemDimensions
        Implements IKeyTip
        Implements ILabel
        Implements IShowLabel
        Implements IScreenTip
        Implements ISuperTip
        Implements ISelect
        Implements ISize
        Implements IAttributeSource
        Implements IRowsAndColumns

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeSelectionChangeEventManager As EventManager(Of BeforeSelectionChangeEventArgs)
        Private ReadOnly _selectionChangedEventManager As EventManager(Of SelectionChangeEventArgs)

        Public Sub New(Optional config As Action(Of IGalleryBuilder) = Nothing,
                       Optional buttons As ICollection(Of IButton) = Nothing,
                       Optional template As RibbonElement = Nothing,
                       Optional tag As Object = Nothing)
            MyBase.New(New List(Of Item), tag)

            Dim builder As GalleryBuilder = New GalleryBuilder(template)

            If config IsNot Nothing Then
                config.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            Me.Buttons = If(buttons, Array.Empty(Of IButton))

            _beforeSelectionChangeEventManager = New EventManager(Of BeforeSelectionChangeEventArgs)($"{NameOf(Gallery)}.{NameOf(BeforeSelectionChange)}")
            _selectionChangedEventManager = New EventManager(Of SelectionChangeEventArgs)($"{NameOf(Gallery)}.{NameOf(SelectionChanged)}")
        End Sub

        Public ReadOnly Property Buttons As ICollection(Of IButton)

#Region "Events"

        Public Custom Event BeforeSelectionChange As EventHandler(Of BeforeSelectionChangeEventArgs)
            AddHandler(value As EventHandler(Of BeforeSelectionChangeEventArgs))
                _beforeSelectionChangeEventManager.Add(_attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of BeforeSelectionChangeEventArgs))
                _beforeSelectionChangeEventManager.Remove(_attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As BeforeSelectionChangeEventArgs)
                _beforeSelectionChangeEventManager.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

        Public Custom Event SelectionChanged As EventHandler(Of SelectionChangeEventArgs)
            AddHandler(value As EventHandler(Of SelectionChangeEventArgs))
                _selectionChangedEventManager.Add(_attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of SelectionChangeEventArgs))
                _selectionChangedEventManager.Remove(_attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As SelectionChangeEventArgs)
                _selectionChangedEventManager.FireEvent(_attributes, sender, e)
            End RaiseEvent
        End Event

        'TODO: Should Raise Click Events

#End Region

#Region "RibbonElement Overrides"

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Get(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

                If Buttons IsNot Nothing AndAlso Buttons.Any() Then
                    Return String.Join(Environment.NewLine, $"<gallery { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Buttons), String.Empty), "</gallery>")
                Else
                    Return $"<gallery { _attributes }/>"
                End If
            End Get
        End Property

#End Region

#Region "Container Overrides"

        Protected Overrides Sub AfterAdd(item As Item)
            AddHandler item.ValueChanged, AddressOf OnChildItemChange

            RefreshNeeded()
        End Sub

        Protected Overrides Sub BeforeClear(currentItems As ICollection(Of Item))
            For Each item As Item In currentItems
                RemoveHandler item.ValueChanged, AddressOf OnChildItemChange
            Next
        End Sub

        Protected Overrides Sub AfterRemove(item As Item)
            RemoveHandler item.ValueChanged, AddressOf OnChildItemChange

            RefreshNeeded()
        End Sub

        Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
            results.Add(Me)

            For Each button As Button In Buttons
                results.Add(button)
            Next
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

#Region "DisplayText"

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Get(Of String)(AttributeCategory.Description)
            End Get
            Set
                _attributes.Set(Of String)(Value, AttributeCategory.Description)
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

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.LabelVisibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.LabelVisibility)
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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Get(Of KeyTip)(AttributeCategory.KeyTip)
            End Get
            Set
                _attributes.Set(Value)
            End Set
        End Property

#End Region

#Region "Item Height And Width"

        Public Property ItemHeight As Integer Implements IItemDimensions.ItemHeight
            Get
                Return _attributes.Get(Of Integer)(AttributeCategory.ItemHeight)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.ItemHeight)
            End Set
        End Property

        Public Property ItemWidth As Integer Implements IItemDimensions.ItemWidth
            Get
                Return _attributes.Get(Of Integer)(AttributeCategory.ItemWidth)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.ItemWidth)
            End Set
        End Property

#End Region

#Region "Image and ShowImage"

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Get(Of RibbonImage)()
            End Get
            Set
                _attributes.Set(Value)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.ImageVisibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.ImageVisibility)
            End Set
        End Property

#End Region

#Region "Action"

        Public Property Selected As Item Implements ISelect.Selected
            Get
                Dim value As Item = _attributes.Get(Of Item)(AttributeCategory.SelectedItemPosition)

                If value IsNot Nothing Then
                    Return value
                Else
                    If Items.Any() Then
                        Return Items.First()
                    Else
                        Return Item.Blank
                    End If
                End If
            End Get
            Set
                Using updateBlock As IDisposable = RefreshSuspension()

                    Dim e As BeforeSelectionChangeEventArgs = New BeforeSelectionChangeEventArgs(Items.ToArray(), Selected, Value)

                    RaiseEvent BeforeSelectionChange(Me, e)

                    If Not e.IsCancelled Then
                        _attributes.Set(e.NewSelection, AttributeCategory.SelectedItemPosition)

                        RaiseEvent SelectionChanged(Me, New SelectionChangeEventArgs(e.NewSelection))
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

#Region "Size"

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Get(Of ControlSize)(AttributeCategory.Size)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Size)
            End Set
        End Property

#End Region

        Public ReadOnly Property Rows As Integer Implements IRowsAndColumns.Rows
            Get
                Return _attributes.Get(Of Integer)(AttributeCategory.Rows)
            End Get
        End Property

        Public ReadOnly Property Columns As Integer Implements IRowsAndColumns.Columns
            Get
                Return _attributes.Get(Of Integer)(AttributeCategory.Columns)
            End Get
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace