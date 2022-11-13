Imports System.Text.RegularExpressions
Imports RibbonX.Api
Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties

Namespace Controls

    Friend NotInheritable Class ButtonGroup
        Inherits Container(Of RibbonElement)
        Implements IButtonGroup
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

        Public Overrides ReadOnly Property XML As String
            Get
                Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

                Return String.Join(Environment.NewLine, $"<buttonGroup { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Me), String.Empty), "</buttonGroup>")
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

        Private ReadOnly Property IReadOnlyCollection_Count As Integer Implements IReadOnlyCollection(Of IRibbonElement).Count
            Get
                Throw New NotImplementedException()
            End Get
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator(Of IRibbonElement) Implements IEnumerable(Of IRibbonElement).GetEnumerator
            Throw New NotImplementedException()
        End Function
    End Class

End Namespace