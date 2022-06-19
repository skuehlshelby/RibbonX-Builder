Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Public NotInheritable Class Box
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IBoxStyle
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional config As Action(Of IBoxBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional items As ICollection(Of RibbonElement) = Nothing, Optional tag As Object = Nothing)
            MyBase.New(items, tag)

            Dim builder As BoxBuilder = New BoxBuilder(template)

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
                Return _
                    String.Join(Environment.NewLine, $"<box { _attributes }>",
                                String.Join(Environment.NewLine, Items), $"</box>")
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

        Public ReadOnly Property BoxStyle As BoxStyle Implements IBoxStyle.BoxStyle
            Get
                Return _attributes.Read(Of BoxStyle)()
            End Get
        End Property

        Public Shared Function Horizontal(ParamArray items() As RibbonElement) As Box
            Return New Box(Sub(bb) bb.Horizontal().Visible(), items:=items)
        End Function

        Public Shared Function Vertical(ParamArray items() As RibbonElement) As Box
            Return New Box(Sub(bb) bb.Vertical().Visible(), items:=items)
        End Function

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace