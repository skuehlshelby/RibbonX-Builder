
Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class Tab
        Inherits ReadOnlyContainer(Of IGroup)
        Implements ITab

        Public Sub New(attributes As IPropertyCollection, items As ICollection(Of IGroup), Optional tag As Object = Nothing)
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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return Attributes.Get(Category.KeyTip)
            End Get
            Set
                Attributes.Set(Value, Category.KeyTip)
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

        Public Overrides Function Clone() As Object
            Return New Tab(CType(Attributes.Clone(), IPropertyCollection), Items.Select(Function(i) i.Clone()).OfType(Of IGroup)().ToArray(), Tag)
        End Function

    End Class

End Namespace