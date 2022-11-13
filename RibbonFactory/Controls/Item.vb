Imports System.Drawing
Imports RibbonX.Api
Imports RibbonX.Builders
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.Properties
Imports RibbonX.Images

Namespace Controls

    ''' <summary>
    ''' Represents an item in a combobox, dropdown, or gallery.
    ''' </summary>
    Friend NotInheritable Class Item
        Inherits RibbonElement
        Implements IItem
        Implements IAttributeSource

        Public Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

        Public Overrides Function ConvertToXml() As String
            Return String.Empty
        End Function

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Get(Of String)(AttributeCategory.Label)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.Label)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Get(Of String)(AttributeCategory.ScreenTip)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return Attributes.Get(Of String)(AttributeCategory.SuperTip)
            End Get
            Set
                Attributes.Set(Value, AttributeCategory.SuperTip)
            End Set
        End Property

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return Attributes.Get(Of RibbonImage)()
            End Get
            Set
                Attributes.Set(Value)
            End Set
        End Property

        Private Function GetAttributes() As AttributeSet Implements IAttributeSource.GetAttributes
            Return Attributes
        End Function

        Public Overrides Function Clone() As Object
            Return New Item(Attributes.Clone(), Tag)
        End Function

        Public Shared ReadOnly Property Blank As IItem
            Get
                Return New Item(Sub(builder) builder.WithId("item-1").WithLabel(String.Empty).WithSuperTip(String.Empty).WithImage(New Bitmap(1, 1)))
            End Get
        End Property
    End Class

End Namespace