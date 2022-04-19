﻿Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls.Events
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public Class EditBox
        Inherits RibbonElement
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IShowLabel
        Implements IKeyTip
        Implements IImage
        Implements IShowImage
        Implements IText
        Implements IMaxLength
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet
        Private ReadOnly _beforeTextChangeHelper As EventManager(Of BeforeTextChangeEventArgs)
        Private ReadOnly _textChangedHelper As EventManager(Of TextChangeEventArgs)

        Public Sub New(steps As Action(Of IEditBoxBuilder), Optional tag As Object = Nothing)
            Me.New(Nothing, steps, tag)
        End Sub

        Public Sub New(template As RibbonElement, steps As Action(Of IEditBoxBuilder), Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As EditBoxBuilder = If(template Is Nothing, New EditBoxBuilder(), New EditBoxBuilder(template))

            steps.Invoke(builder)

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded

            _beforeTextChangeHelper = New EventManager(Of BeforeTextChangeEventArgs)($"{NameOf(EditBox)}.{NameOf(BeforeTextChangeEventArgs)}")
            _textChangedHelper = New EventManager(Of TextChangeEventArgs)($"{NameOf(EditBox)}.{NameOf(TextChangeEventArgs)}")
        End Sub

#Region "Events"

        Public Custom Event BeforeTextChange As EventHandler(Of BeforeTextChangeEventArgs)
            AddHandler(value As EventHandler(Of BeforeTextChangeEventArgs))
                _beforeTextChangeHelper.Add(_attributes, value)
            End AddHandler

            RemoveHandler(value As EventHandler(Of BeforeTextChangeEventArgs))
                _beforeTextChangeHelper.Remove(_attributes, value)
            End RemoveHandler

            RaiseEvent(sender As Object, e As BeforeTextChangeEventArgs)
                _beforeTextChangeHelper.FireEvent(_attributes, sender, e)
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

#Region "Base Class Overrides"

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Read(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<editBox { _attributes }/>"
            End Get
        End Property

#End Region

#Region "Display Text"

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Read(Of KeyTip)(AttributeCategory.KeyTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.KeyTip)
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

#Region "Image"

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