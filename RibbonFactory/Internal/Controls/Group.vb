
Imports RibbonX.Controls.Base
Imports RibbonX.Images
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class Group
        Inherits ReadOnlyContainer(Of IRibbonElement)
        Implements IGroup

        Public Sub New(attributes As IPropertyCollection, items As ICollection(Of IRibbonElement), Optional tag As Object = Nothing)
            MyBase.New(attributes, items, tag)
        End Sub

        Public Property Label As String Implements ILabel.Label
            Get
                Return Attributes.Get(Category.Label)
            End Get
            Set
                Attributes.Set(Value, Category.Label)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return Attributes.Get(Category.Visibility)
            End Get
            Set
                Attributes.Set(Value, Category.Visibility)
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

        Public Overrides Function Clone() As Object
            Return New Group(CType(Attributes.Clone(), IPropertyCollection), Items.Select(Function(i) i.Clone()).OfType(Of IRibbonElement)().ToArray(), Tag)
        End Function

    End Class

End Namespace