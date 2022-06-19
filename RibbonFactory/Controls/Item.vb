Imports System.Drawing
Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Images

Namespace Controls

    ''' <summary>
    ''' Represents an item in a combobox, dropdown, or gallery.
    ''' </summary>
    Public NotInheritable Class Item
        Inherits RibbonElement
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IImage
        Implements IAttributeSource

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(Optional config As Action(Of IItemBuilder) = Nothing, Optional template As RibbonElement = Nothing, Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As ItemBuilder = New ItemBuilder(template)

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
                Return String.Empty
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

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return _attributes.Read(Of RibbonImage)()
            End Get
            Set
                _attributes.Write(Value)
            End Set
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return _attributes
        End Function

        Public Shared ReadOnly Property Blank As Item = New Item(Sub(builder) builder.WithId("item-1").WithLabel(String.Empty).WithSuperTip(String.Empty).WithImage(New Bitmap(1, 1)))

    End Class

End Namespace