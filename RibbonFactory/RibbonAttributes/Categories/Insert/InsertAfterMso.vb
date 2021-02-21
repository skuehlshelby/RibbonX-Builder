Namespace RibbonAttributes.Categories.Insert

    Friend Class InsertAfterMso
        Inherits RibbonAttribute(Of String, InsertAfterMso)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(InsertAfterMso), GetValue())
            End Get
        End Property
    End Class
End NameSpace