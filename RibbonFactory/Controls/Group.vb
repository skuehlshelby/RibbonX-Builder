Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Images
Imports RibbonX.SimpleTypes

Namespace Controls

    Public Class Group
        Inherits Container(Of RibbonElement)
        Implements ILabel
        Implements IVisible
        Implements IKeyTip
        Implements IImage
        Implements IScreenTip
        Implements ISuperTip
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional config As Action(Of IGroupBuilder) = Nothing,
                       Optional controls As ICollection(Of RibbonElement) = Nothing,
                       Optional template As RibbonElement = Nothing,
                       Optional tag As Object = Nothing)
            MyBase.New(controls, tag)

            Dim builder As GroupBuilder = New GroupBuilder(template)

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
                Return String.Join(Environment.NewLine, $"<group { _attributes }>", String.Join(Environment.NewLine, Items), $"</group>")
            End Get
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Get(Of String)(AttributeCategory.Label)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Get(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Get(Of KeyTip)()
            End Get
            Set
                _attributes.Set(Value)
            End Set
        End Property

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Get(Of RibbonImage)()
            End Get
            Set
                _attributes.Set(Value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Get(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Get(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Set(Value, AttributeCategory.SuperTip)
            End Set
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes.Clone()
        End Function

    End Class

End Namespace