Imports System.Text.RegularExpressions
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public NotInheritable Class ButtonGroup
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional tag As Object = Nothing)
            Me.New(Nothing, Nothing, Nothing, tag)
        End Sub

        Public Sub New(steps As Action(Of IButtonGroupBuilder), Optional tag As Object = Nothing)
            Me.New(steps, Nothing, Nothing, tag)
        End Sub

        Public Sub New(items As ButtonGroupControls, Optional tag As Object = Nothing)
            Me.New(Nothing, items, Nothing, tag)
        End Sub

        Public Sub New(steps As Action(Of IButtonGroupBuilder), items As ButtonGroupControls, Optional tag As Object = Nothing)
            Me.New(steps, items, Nothing, tag)
        End Sub

        Public Sub New(steps As Action(Of IButtonGroupBuilder), items As ButtonGroupControls, template As RibbonElement, Optional tag As Object = Nothing)
            MyBase.New(If(items Is Nothing, Array.Empty(Of RibbonElement)(), items.ToArray()), tag)

            Dim builder As ButtonGroupBuilder = New ButtonGroupBuilder(template)

            If steps IsNot Nothing Then
                steps.Invoke(builder)
            End If

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
                Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

                Return String.Join(Environment.NewLine, $"<buttonGroup { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Me), String.Empty), "</buttonGroup>")
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

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace