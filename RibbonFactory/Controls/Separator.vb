Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Controls

    Public NotInheritable Class Separator
        Inherits RibbonElement
        Implements IVisible
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional tag As Object = Nothing)
            Me.New(Nothing, Nothing, tag)
        End Sub

        Public Sub New(configuration As Action(Of ISeparatorBuilder), Optional tag As Object = Nothing)
            Me.New(Nothing, configuration, tag)
        End Sub

        Public Sub New(template As RibbonElement, configuration As Action(Of ISeparatorBuilder), Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As SeparatorBuilder = If(template Is Nothing, New SeparatorBuilder(), New SeparatorBuilder(template))

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
                Return $"<separator { _attributes }/>"
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