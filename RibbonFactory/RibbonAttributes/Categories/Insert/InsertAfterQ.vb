Namespace RibbonAttributes.Categories.Insert
    Public Class InsertAfterQ
        Inherits RibbonAttribute(Of String, InsertAfterQ)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(InsertAfterQ), GetValue())
            End Get
        End Property
    End Class
End NameSpace