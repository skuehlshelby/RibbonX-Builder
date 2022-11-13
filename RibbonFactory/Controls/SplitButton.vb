Imports System.Text.RegularExpressions
Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Public NotInheritable Class SplitButton
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IEnabled
        Implements IKeyTip
        Implements ISize
        Implements IShowLabel
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(button As IButton, menu As Menu, Optional config As Action(Of ISplitButtonBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            Me.New(DirectCast(button, RibbonElement), menu, config, template, tag)
        End Sub

        Public Sub New(toggleButton As ToggleButton, menu As Menu, Optional config As Action(Of ISplitButtonBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            Me.New(DirectCast(toggleButton, RibbonElement), menu, config, template, tag)
        End Sub

        Private Sub New(button As RibbonElement, menu As Menu, Optional config As Action(Of ISplitButtonBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            MyBase.New(New RibbonElement() {button, menu}, tag)

            Dim builder As SplitButtonBuilder = New SplitButtonBuilder(template)

            If config IsNot Nothing Then
                config.Invoke(builder)
            End If

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Get(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Dim regex As Regex = New Regex("(?:size|getSize|visible|getVisible)=""\w+""")

                Return String.Join(Environment.NewLine, $"<splitButton { _attributes }>",
                                    regex.Replace(String.Join(Environment.NewLine, Items), String.Empty),
                                    "</splitButton>")
            End Get
        End Property

        Public ReadOnly Property Button As RibbonElement
            Get
                Return Items(0)
            End Get
        End Property

        Public ReadOnly Property Menu As Menu
            Get
                Return DirectCast(Items(1), Menu)
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Enabled)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Get(Of KeyTip)()
            End Get
            Set
                _attributes.Set(Value)
            End Set
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Get(Of ControlSize)(AttributeCategory.Size)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Size)
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

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace