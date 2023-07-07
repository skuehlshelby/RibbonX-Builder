
Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi
Imports RibbonX.Properties

Namespace Controls

    Friend NotInheritable Class Separator
        Inherits RibbonElement
        Implements ISeparator

        Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
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
            Return New Separator(CType(Attributes.Clone(), IPropertyCollection), Tag)
        End Function

    End Class

End Namespace