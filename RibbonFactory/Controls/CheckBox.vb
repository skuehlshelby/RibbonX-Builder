﻿Imports RibbonX.Builders
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Utilities
Imports RibbonX.SimpleTypes

Namespace Controls

    Public NotInheritable Class CheckBox
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IDescription
        Implements IKeyTip
        Implements IPressed
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeToggleManager As EventManager(Of BeforeToggleChangeEventArgs)
        Private ReadOnly _onToggleManager As EventManager(Of ToggleChangeEventArgs)

        Public Sub New(Optional config As Action(Of ICheckBoxBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As CheckBoxBuilder = New CheckBoxBuilder(template)

            If config IsNot Nothing Then
                config.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            _beforeToggleManager = New EventManager(Of BeforeToggleChangeEventArgs)($"{NameOf(CheckBox)}.{NameOf(BeforeStateChange)}")
            _onToggleManager = New EventManager(Of ToggleChangeEventArgs)($"{NameOf(CheckBox)}.{NameOf(StateChanged)}")
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
                Return $"<checkBox { _attributes }/>"
            End Get
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

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.SuperTip)
            End Set
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
                        _attributes.Write(e.FutureState, AttributeCategory.Pressed)

                        RaiseEvent StateChanged(Me, New ToggleChangeEventArgs(e.FutureState))
                    End If
                End Using
            End Set
        End Property

#End Region

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace