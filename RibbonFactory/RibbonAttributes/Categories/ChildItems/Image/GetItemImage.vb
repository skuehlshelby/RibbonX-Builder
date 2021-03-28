Imports stdole

Namespace RibbonAttributes.Categories.ChildItems.Image

    Friend NotInheritable Class GetItemImage
        Inherits RibbonAttribute(Of String, GetItemImage)

        Public Sub New (callback As FromControl(Of IPictureDisp))
            MyBase.New(callback.Method.Name)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, CamelCase(NameOf(GetItemImage)), Value)
            End Get
        End Property

    End Class

End NameSpace