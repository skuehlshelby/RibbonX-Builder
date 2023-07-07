Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi
Imports RibbonX.Properties

Namespace Controls

    Friend NotInheritable Class ButtonGroup
        Inherits ReadOnlyContainer(Of IRibbonElement)
        Implements IButtonGroup

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

        Public Overrides Function Clone() As Object
            Return New ButtonGroup(CType(Attributes.Clone(), IPropertyCollection), Items.Select(Function(i) i.Clone()).OfType(Of IRibbonElement)().ToArray(), Tag)
        End Function

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            Return $"<buttonGroup {Attributes.ToXml(exclude)}>{String.Join(Environment.NewLine, Items.Select(Function(btn) btn.ToXml(ExcludedAttributes.Size)))}</buttonGroup>"
        End Function

    End Class

End Namespace