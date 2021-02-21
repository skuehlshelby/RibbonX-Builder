Namespace RibbonAttributes.Categories.Insert
    Public Class InsertBeforeQ
        Inherits RibbonAttribute(Of String, InsertBeforeQ)

        Public Sub New(Value As String)
            MyBase.New(Value)
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(InsertBeforeQ), GetValue())
            End Get
        End Property
    End Class
End NameSpace