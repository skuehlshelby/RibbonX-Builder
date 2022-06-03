Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Builders
Imports RibbonX.ControlInterfaces
Imports RibbonX.RibbonAttributes

Namespace Controls

    Public Class Group
        Inherits Container(Of RibbonElement)
        Implements ILabel
        Implements IVisible
        Implements IKeyTip
        Implements IImage
        Implements IScreenTip
        Implements ISuperTip
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(ParamArray controls() As RibbonElement)
            Me.New(Nothing, controls, Nothing, Nothing)
        End Sub

        Public Sub New(configuration As Action(Of IGroupBuilder), controls As ICollection(Of RibbonElement), Optional tag As Object = Nothing)
            Me.New(configuration, controls, Nothing, tag)
        End Sub

        Public Sub New(configuration As Action(Of IGroupBuilder), controls As ICollection(Of RibbonElement), template As RibbonElement, Optional tag As Object = Nothing)
            MyBase.New(controls, tag)

            Dim builder As GroupBuilder = If(template Is Nothing, New GroupBuilder(), New GroupBuilder(template))

            If configuration IsNot Nothing Then
                configuration.Invoke(builder)
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
                If Items.Any() Then
                    Return String.Join(Environment.NewLine, $"<group { _attributes }>", String.Join(Environment.NewLine, Items), $"</group>")
                Else
                    Return $"<group { _attributes } />"
                End If
            End Get
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Read(Of String)(AttributeCategory.Label)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Read(Of Boolean)(AttributeCategory.Visibility)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.Visibility)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Read(Of KeyTip)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Read(Of RibbonImage)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Read(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                _attributes.Write(Value, AttributeCategory.SuperTip)
            End Set
        End Property

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes.Clone()
        End Function

    End Class

End Namespace