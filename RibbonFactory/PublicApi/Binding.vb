Imports System.ComponentModel

Namespace Api

    Public Interface IBinding

    End Interface

    Friend NotInheritable Class Binding
        Implements IBinding

        Private ReadOnly source As Object
        Private ReadOnly sourceProperty As String
        Private ReadOnly target As Object
        Private ReadOnly targetProperty As String

        Private Sub OnPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            If sender Is source AndAlso sourceProperty Is e.PropertyName Then
            ElseIf sender Is target AndAlso targetProperty Is e.PropertyName Then

            End If
        End Sub
    End Class

End Namespace