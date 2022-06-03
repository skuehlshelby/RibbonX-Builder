Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Builders
Imports RibbonX.ControlInterfaces
Imports RibbonX.RibbonAttributes

Namespace Controls

    Public NotInheritable Class MenuSeparator
        Inherits RibbonElement
        Implements ITitle
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(configuration As Action(Of IMenuSeparatorBuilder), Optional tag As Object = Nothing)
            Me.New(Nothing, configuration, tag)
        End Sub

        Public Sub New(template As RibbonElement, configuration As Action(Of IMenuSeparatorBuilder), Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As MenuSeparatorBuilder = If(template Is Nothing, New MenuSeparatorBuilder(), New MenuSeparatorBuilder(template))

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
                Return $"<menuSeparator { _attributes }/>"
            End Get
        End Property

        Public Property Title As String Implements ITitle.Title
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Title)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Title)
            End Set
        End Property

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace