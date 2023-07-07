
Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class SplitButton
        Inherits ReadOnlyContainer(Of IRibbonElement)
        Implements ISplitButton

        Public Sub New(attributes As IPropertyCollection, button As IRibbonElement, menu As IMenu, Optional tag As Object = Nothing)
            MyBase.New(attributes, {Utilities.NotNull(button), Utilities.NotNull(menu)}, tag)
        End Sub

        Public ReadOnly Property Button As IRibbonElement Implements ISplitButton.Button
            Get
                Return Items(0)
            End Get
        End Property

        Public ReadOnly Property Menu As IMenu Implements ISplitButton.Menu
            Get
                Return DirectCast(Items(1), IMenu)
            End Get
        End Property

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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Category.KeyTip)
            End Get
            Set
                Attributes.Set(Value, Category.KeyTip)
            End Set
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return Attributes.Get(Category.Size)
            End Get
            Set
                Attributes.Set(Value, Category.Size)
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

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            Return MyBase.ToXml(ExcludedAttributes.SizeAndVisibility)
        End Function

        Public Overrides Function Clone() As Object
            Return New SplitButton(CType(Attributes.Clone(), IPropertyCollection), CType(Button.Clone(), IRibbonElement), CType(Menu.Clone(), IMenu), Tag)
        End Function

    End Class

End Namespace