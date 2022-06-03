Imports System.Text.RegularExpressions
Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Builders
Imports RibbonX.ControlInterfaces
Imports RibbonX.Enums
Imports RibbonX.RibbonAttributes

Namespace Controls

    Public NotInheritable Class SplitButton
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IEnabled
        Implements IKeyTip
        Implements ISize
        Implements IShowLabel
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(configuration As Action(Of ISplitButtonBuilder), button As Button, menu As Menu, Optional tag As Object = Nothing)
            Me.New(Nothing, configuration, DirectCast(button, RibbonElement), menu, tag)
        End Sub

        Public Sub New(configuration As Action(Of ISplitButtonBuilder), button As ToggleButton, menu As Menu, Optional tag As Object = Nothing)
            Me.New(Nothing, configuration, DirectCast(button, RibbonElement), menu, tag)
        End Sub

        Public Sub New(template As RibbonElement, configuration As Action(Of ISplitButtonBuilder), button As Button, menu As Menu, Optional tag As Object = Nothing)
            Me.New(template, configuration, DirectCast(button, RibbonElement), menu, tag)
        End Sub

        Public Sub New(template As RibbonElement, configuration As Action(Of ISplitButtonBuilder), button As ToggleButton, menu As Menu, Optional tag As Object = Nothing)
            Me.New(template, configuration, DirectCast(button, RibbonElement), menu, tag)
        End Sub

        Private Sub New(template As RibbonElement, configuration As Action(Of ISplitButtonBuilder), button As RibbonElement, menu As Menu, Optional tag As Object = Nothing)
            MyBase.New(New RibbonElement() {button, menu}, tag)

            Dim builder As SplitButtonBuilder = If(template Is Nothing, New SplitButtonBuilder(), New SplitButtonBuilder(template))

            configuration.Invoke(builder)

            _attributes = builder.Build()

            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Read(Of String)(AttributeCategory.IdType)
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Dim regex As Regex = New Regex("(?:size|getSize|visible|getVisible)=""\w+""")

                If Items.Any() Then
                    Return _
                        String.Join(Environment.NewLine, $"<splitButton { _attributes }>",
                                    regex.Replace(String.Join(Environment.NewLine, Items), String.Empty),
                                    "</splitButton>")

                Else
                    Return $"<splitButton { _attributes } />"
                End If
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
                Return _attributes.Read(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Enabled)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Enabled)
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

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Read(Of ControlSize)(AttributeCategory.Size)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Size)
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

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace