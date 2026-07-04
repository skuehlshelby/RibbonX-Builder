
Imports RibbonX.Controls.Base
Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties

Namespace Controls

    ''' <summary>
    ''' Represents an item in a combobox, dropdown, or gallery.
    ''' </summary>
    Friend NotInheritable Class Item
        Inherits RibbonElement
        Implements IItem

        Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Get(Category.Label)
            End Get
            Set
                Attributes.Set(Value, Category.Label)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return Attributes.Get(Category.ScreenTip)
            End Get
            Set
                Attributes.Set(Value, Category.ScreenTip)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return Attributes.Get(Category.SuperTip)
            End Get
            Set
                Attributes.Set(Value, Category.SuperTip)
            End Set
        End Property

        Public Property Image As RibbonImage Implements IImage.Image
            Get
                Return Attributes.Get(Category.Image)
            End Get
            Set
                Attributes.Set(Value, Category.Image)
            End Set
        End Property

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            Return String.Empty
        End Function

        Public Overrides Function Clone() As Object
            Return New Item(CType(Attributes.Clone(), IPropertyCollection), Tag)
        End Function

    End Class

End Namespace