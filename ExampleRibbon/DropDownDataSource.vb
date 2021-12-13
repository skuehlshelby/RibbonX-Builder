Imports System.Collections.ObjectModel
Imports System.Drawing
Imports System.IO
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Controls
Imports stdole

Public Module DropDownDataSource
    
    Public Function GetFilesOnDesktop() As IEnumerable(Of FileInfo)
        Dim results As ICollection(Of FileInfo) = New Collection(Of FileInfo)

        For Each path As String In Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
            results.Add(New FileInfo(path))
        Next

        Return results
    End Function

    Public Function ConvertFileToDropDownItem(file As FIleInfo, callback As FromControl(Of IPictureDisp)) As Item
        Return New ItemBuilder().
            WithLabel(file.Name).
            WithSuperTip(file.FullName).
            WithImage(Icon.ExtractAssociatedIcon(file.FullName), callback).
            Build()
    End Function

    Public Sub OpenFile(path As String)
        Try
            Dim file As FileInfo = New FileInfo(path)

            If file.Exists Then
                Using process As Process = New Process()
                    process.StartInfo.FileName = file.FullName
                    process.Start()
                End Using
            End If
        Catch ex As Exception

        End Try
    End Sub

End Module