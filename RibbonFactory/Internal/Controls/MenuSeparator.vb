Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi
Imports RibbonX.Properties

Namespace Controls

    Friend NotInheritable Class MenuSeparator
        Inherits RibbonElement
        Implements IMenuSeparator

        Public Sub New(attributes As IPropertyCollection, Optional tag As Object = Nothing)
            MyBase.New(attributes, tag)
        End Sub

        Public Property Title As String Implements ITitle.Title
            Get
                Return Attributes.Get(Category.Title)
            End Get
            Set
                Attributes.Set(Value, Category.Title)
            End Set
        End Property

        Public Overrides Function Clone() As Object
            Return New MenuSeparator(CType(Attributes.Clone(), IPropertyCollection), Tag)
        End Function

    End Class

End Namespace