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
        Public Function NoDoubleSpace(input As String) As String
            If String.IsNullOrWhiteSpace(input) Then
                Return String.Empty
            Else
                Dim output(input.Length) As Char
                Dim skipped As Boolean = False
                Dim current As Integer = 0

                For Each letter As Char In input.ToCharArray()
                    If Char.IsWhiteSpace(letter) Then
                        If Not skipped Then
                            If 0 < current Then
                                output(current) = " "c
                                current += 1
                            End If

                            skipped = True
                        End If
                    Else
                        skipped = False
                        output(current) = letter
                        current += 1
                    End If
                Next

                Return New String(output, 0, current)
            End If
        End Function

    End Module
End NameSpace