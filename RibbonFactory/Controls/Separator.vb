Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties

Namespace Controls

    Public NotInheritable Class Separator
        Inherits RibbonElement
        Implements IVisible
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional config As Action(Of ISeparatorBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As SeparatorBuilder = New SeparatorBuilder(template)

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

        Public Shared Function CreateVisible(Optional additionalConfiguration As Action(Of ISeparatorBuilder) = Nothing) As Separator
            Dim invisibleAction As Action(Of ISeparatorBuilder) = Sub(sb) sb.Visible()
            Dim configuration As Action(Of ISeparatorBuilder) = DirectCast([Delegate].Combine(invisibleAction, additionalConfiguration), Action(Of ISeparatorBuilder))

            Return New Separator(configuration)
        End Function

        Public Shared Function CreateInvisible(Optional additionalConfiguration As Action(Of ISeparatorBuilder) = Nothing) As Separator
            Dim invisibleAction As Action(Of ISeparatorBuilder) = Sub(sb) sb.Invisible()
            Dim configuration As Action(Of ISeparatorBuilder) = DirectCast([Delegate].Combine(invisibleAction, additionalConfiguration), Action(Of ISeparatorBuilder))

            Return New Separator(configuration)
        End Function

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace