
Imports RibbonX.Controls.Base
Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class Menu
        Inherits ReadOnlyContainer(Of IRibbonElement)
        Implements IMenu

        Public Sub New(attributes As IPropertyCollection, items As ICollection(Of IRibbonElement), Optional tag As Object = Nothing)
            MyBase.New(attributes, items, tag)
        End Sub

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Get(Category.Visibility)
            End Get
            Set
                Attributes.Set(Value, Category.Visibility)
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return Attributes.Get(Category.Enabled)
            End Get
            Set
                Attributes.Set(Value, Category.Enabled)
            End Set
        End Property

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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Category.KeyTip)
            End Get
            Set
                Attributes.Set(Value, Category.KeyTip)
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

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return Attributes.Get(Category.ImageVisibility)
            End Get
            Set
                Attributes.Set(Value, Category.ImageVisibility)
            End Set
        End Property

        Public Property Description As String Implements IDescription.Description
            Get
                Return Attributes.Get(Category.Description)
            End Get
            Set
                Attributes.Set(Value, Category.Description)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return Attributes.Get(Category.LabelVisibility)
            End Get
            Set
                Attributes.Set(Value, Category.LabelVisibility)
            End Set
        End Property

        Public ReadOnly Property ItemSize As ControlSize Implements IItemSize.ItemSize
            Get
                Return Attributes.Get(Category.ItemSize)
            End Get
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return Attributes.Get(Category.Size)
            End Get
            Set
                Attributes.Set(Value, Category.Size)
            End Set
        End Property

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            Return MyBase.ToXml(ExcludedAttributes.SizeAndVisibility)
        End Function

        Public Overrides Function Clone() As Object
            Return New Menu(CType(Attributes.Clone(), IPropertyCollection), Items.Select(Function(i) i.Clone()).OfType(Of IRibbonElement)().ToArray(), Tag)
        End Function

    End Class

End Namespace