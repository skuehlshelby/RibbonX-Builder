Namespace RibbonAttributes.Categories.Title
    
    Public Class GetTitle
        Inherits Title
        
        Private ReadOnly _callback As String

        Public Sub New(value As String, callback As FromControl(Of String))
            MyBase.New(value)
            _callback = callback.Method.Name
        End Sub

        Public Sub SetValue(newValue As String)
            If Not Value.Equals(newValue) Then
                Value = newValue
            End If
        End Sub

        Public Overrides ReadOnly Property XML As String
            Get
                Return String.Format(XML_TEMPLATE, NameOf(GetTitle).CamelCase(), _callback)
            End Get
        End Property
    End Class
    
End NameSpace