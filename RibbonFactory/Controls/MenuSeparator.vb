Imports RibbonX.Builders
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Controls.Attributes

Namespace Controls

    Public NotInheritable Class MenuSeparator
        Inherits RibbonElement
        Implements ITitle
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional config As Action(Of IMenuSeparatorBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As MenuSeparatorBuilder = New MenuSeparatorBuilder(template)

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
                Return $"<menuSeparator { _attributes }/>"
            End Get
        End Property

        Public Property Title As String Implements ITitle.Title
            Get
                Return _attributes.Get(Of String)(AttributeCategory.Title)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Title)
            End Set
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace