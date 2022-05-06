Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

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
        Implements IDefaultProvider

        Private ReadOnly _attributes As AttributeSet

        Public Sub New(steps As Action(Of IItemBuilder), Optional tag As Object = Nothing)
            Me.New(steps, Nothing, tag)
        End Sub

        Public Sub New(steps As Action(Of IItemBuilder), template As RibbonElement, Optional tag As Object = Nothing)
            MyBase.New(tag)

            Dim builder As ItemBuilder = New ItemBuilder(template)

            If steps IsNot Nothing Then
                steps.Invoke(builder)
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

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes
        End Function

        Public Shared ReadOnly Property Blank As Item = New Item(Sub(builder) builder.WithId("item-1").WithLabel(String.Empty).WithSuperTip(String.Empty).WithImage(New Bitmap(1, 1)))

    End Class

End Namespace