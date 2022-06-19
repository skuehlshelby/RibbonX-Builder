Imports System.Text.RegularExpressions
Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties

Namespace Controls

    Public NotInheritable Class ButtonGroup
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional config As Action(Of IButtonGroupBuilder) = Nothing,
                       Optional items As ButtonGroupControls = Nothing,
                       Optional template As RibbonElement = Nothing,
                       Optional tag As Object = Nothing)
            MyBase.New(If(items Is Nothing, Array.Empty(Of RibbonElement)(), items.ToArray()), tag)

            Dim builder As ButtonGroupBuilder = New ButtonGroupBuilder(template)

            If config IsNot Nothing Then
                config.Invoke(builder)
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

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace