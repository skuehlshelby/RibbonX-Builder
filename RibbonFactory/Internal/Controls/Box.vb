Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace Controls

    Friend NotInheritable Class Box
        Inherits ReadOnlyContainer(Of IRibbonElement)
        Implements IBox

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

        Public ReadOnly Property BoxStyle As BoxStyle Implements IBoxStyle.BoxStyle
            Get
                Return Attributes.Get(Category.BoxStyle)
            End Get
        End Property

        Public Overrides Function Clone() As Object
            Return New Box(CType(Attributes.Clone(), IPropertyCollection), Items.Select(Function(i) i.Clone()).OfType(Of IRibbonElement)().ToArray(), Tag)
        End Function

    End Class

End Namespace