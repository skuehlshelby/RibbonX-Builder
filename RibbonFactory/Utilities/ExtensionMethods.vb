Imports System.Runtime.CompilerServices

Namespace Utilities

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

        <Extension()>
        Public Function IsDerivedFromGenericType(type As Type, generic As Type) As Boolean
            While type IsNot Nothing AndAlso type IsNot GetType(Object)
                Dim check As Type = If(type.IsGenericType, type.GetGenericTypeDefinition(), type)

                If check Is generic Then
                    Return True
                End If

                type = type.BaseType
            End While

            Return False
        End Function
    End Module
End NameSpace