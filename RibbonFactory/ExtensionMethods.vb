Imports System.Runtime.CompilerServices

Friend Module ExtensionMethods
    
    <Extension()>
    Public Function CamelCase(input As Object) As String 
        Return input.ToString().Substring(0, 1).ToLower() & input.ToString().Substring(1)
    End Function

    <Extension()>
    Public Function IndexOf(Of T)(collection As ICollection(Of T), target As T) As Integer
        Dim index As Integer = 0

        For Each item As T In collection
            If item.Equals(target) Then
                Return index
            Else
                index += 1
            End If
        Next

        Return -1
    End Function
End Module