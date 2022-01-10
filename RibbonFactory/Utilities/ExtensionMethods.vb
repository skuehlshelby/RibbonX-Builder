Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Xml
Imports System.Xml.Schema
Imports RibbonFactory.Containers
Imports RibbonFactory.Utilities.Testing

Namespace Utilities

    Friend Module ExtensionMethods

        <Extension>
        Public Function GetErrors(ribbon As Ribbon) As RibbonErrorLog
            Return GetErrors(ribbon.RibbonX)
        End Function

        Public Function GetErrors(ribbonX As String) As RibbonErrorLog
            Return New RibbonErrorLog(ribbonX, GetCustomUIVersion2009())
        End Function

        Public Function GetCustomUIVersion2009() As XmlSchema
            Using reader As StreamReader = New StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("RibbonFactory.RibbonX.xsd"))
                Return XmlSchema.Read(XmlReader.Create(reader), AddressOf ValidationEventHandler)
            End Using
        End Function

        Private Sub ValidationEventHandler(sender As Object, e As ValidationEventArgs)
            Console.WriteLine($"{e.Severity.ToString().ToUpper}: {e.Message}")
        End Sub
    
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

        <Extension()>
        Public Function ImplementsGenericInterface(type As Type, generic As Type) As Boolean
            While type IsNot Nothing AndAlso type IsNot GetType(Object)
                For Each iFace As Type In type.GetInterfaces()
                    Dim check As Type = If(iFace.IsGenericType, iFace.GetGenericTypeDefinition(), iFace)

                    If check Is generic Then
                        Return True
                    End If
                Next

                type = type.BaseType
            End While

            Return False
        End Function

        <Extension()>
        Public Function NextBoolean(random As Random) As Boolean
            Return random.Next Mod 2 = 0
        End Function

    End Module
End NameSpace