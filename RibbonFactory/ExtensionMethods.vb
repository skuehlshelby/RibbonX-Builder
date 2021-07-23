Imports System.Runtime.CompilerServices

Friend Module ExtensionMethods
    
    <Extension()>
    Public Function CamelCase(input As Object) As String 
        Return input.ToString().Substring(0, 1).ToLower() & input.ToString().Substring(1)
    End Function
End Module