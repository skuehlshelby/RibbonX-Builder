Namespace RibbonAttributes.Categories.Insert
    Public Class InsertBeforeMso
        Inherits RibbonAttribute(Of String, InsertBeforeMso)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(InsertBeforeMso), GetValue())
            End Get
        End Property
    End Class
End NameSpace