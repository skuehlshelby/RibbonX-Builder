﻿Imports System.Text.RegularExpressions
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

    Public NotInheritable Class DropDown
        Inherits Container(Of Item)
        Implements ISelect
        Implements IEnabled
        Implements IVisible
        Implements IKeyTip
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IShowLabel
        Implements IImage
        Implements IShowImage
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeSelectionChangeEventManager As EventManager(Of BeforeSelectionChangeEventArgs)
        Private ReadOnly _selectionChangedEventManager As EventManager(Of SelectionChangeEventArgs)

        Public Sub New(Optional config As Action(Of IDropdownBuilder) = Nothing,
                       Optional buttons As ICollection(Of Button) = Nothing,
                       Optional template As RibbonElement = Nothing,
                       Optional tag As Object = Nothing)
            MyBase.New(New List(Of Item)(), tag)

            Dim builder As DropDownBuilder = New DropDownBuilder(template)

            If config IsNot Nothing Then
                config.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            Me.Buttons = If(buttons, Array.Empty(Of Button)())

            _beforeSelectionChangeEventManager = New EventManager(Of BeforeSelectionChangeEventArgs)($"{NameOf(DropDown)}.{NameOf(BeforeSelectionChange)}")
            _selectionChangedEventManager = New EventManager(Of SelectionChangeEventArgs)($"{NameOf(DropDown)}.{NameOf(SelectionChanged)}")
        End Sub

        Public ReadOnly Property Buttons As ICollection(Of Button)

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

#End Region

#Region "RibbonElement Overrides"

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Read(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

                If Buttons IsNot Nothing AndAlso Buttons.Any() Then
                    Return String.Join(Environment.NewLine, $"<dropDown { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Buttons), String.Empty), "</dropDown>")
                Else
                    Return $"<dropDown { _attributes }/>"
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

        Protected Overrides Sub AfterClear()
            RefreshNeeded()
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
                Return _attributes.Read(Of KeyTip)()
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

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.ScreenTip)
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

#Region "Action"

        Public Property Selected As Item Implements ISelect.Selected
            Get
                Dim value As Item = _attributes.Read(Of Item)(AttributeCategory.SelectedItemPosition)

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
                        _attributes.Write(e.NewSelection, AttributeCategory.SelectedItemPosition)

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

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace