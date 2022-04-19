Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public NotInheritable Class Box
        Inherits Container(Of RibbonElement)
        Implements IVisible
        Implements IBoxStyle
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(steps As Action(Of IBoxBuilder), items As ICollection(Of RibbonElement), Optional tag As Object = Nothing)
            Me.New(Nothing, steps, items, tag)
        End Sub

        Public Sub New(template As RibbonElement, steps As Action(Of IBoxBuilder), items As ICollection(Of RibbonElement), Optional tag As Object = Nothing)
            MyBase.New(items, tag)

            Dim builder As BoxBuilder = If(template Is Nothing, New BoxBuilder(), New BoxBuilder(template))

            steps.Invoke(builder)

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
            Return New Box(Sub(bb) bb.Horizontal().Visible(), items)
        End Function

        Public Shared Function Vertical(ParamArray items() As RibbonElement) As Box
            Return New Box(Sub(bb) bb.Vertical().Visible(), items)
        End Function

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function
    End Class

End Namespace