Imports System.Collections.ObjectModel
Imports System.Drawing
Imports System.IO
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports stdole

Public Class DesktopFilesGroup
    Implements IRibbonElementSource

    Private _dropDown As DropDown
    Private _openButton As Button

    Public Function GetElement(callbacks As ICreateCallbacks) As RibbonElement Implements IRibbonElementSource.GetElement
        _dropDown = New DropdownBuilder().
                WithScreenTip("Desktop Files").
                WithSuperTip("The files on your desktop.").
                WithSize("A DropDown This Big").
                GetItemCountFrom(AddressOf callbacks.GetItemCount).
                GetItemIdFrom(AddressOf callbacks.GetItemId).
                GetItemImageFrom(AddressOf callbacks.GetItemImage).
                GetItemLabelFrom(AddressOf callbacks.GetItemLabel).
                GetItemSuperTipFrom(AddressOf callbacks.GetItemSuperTip).
                GetItemScreenTipFrom(AddressOf callbacks.GetItemScreenTip).
                GetSelectedItemIdFrom(AddressOf callbacks.GetSelectedItemId).
                ThatDoes(Sub() Return, AddressOf callbacks.OnSelectionChange).
                Build()

        For Each file As FileInfo In GetFilesOnDesktop()
            _dropDown.Add(ConvertFileToDropDownItem(file, AddressOf callbacks.GetImage))
        Next

        Dim dropDownLabel As LabelControl = New LabelControlBuilder().
                WithLabel("Desktop Files:").
                ShowLabel().
                Build()

        _openButton = New ButtonBuilder().
                WithLabel("Open File").
                WithSuperTip("Open or launch the selected file/program.").
                WithImage(Enums.ImageMSO.Common.FileOpen).
                ThatDoes(AddressOf callbacks.OnAction, Sub() OpenFile(_dropDown.Selected.SuperTip)).
                Build()

        Return New GroupBuilder().
                WithControl(BoxBuilder.Horizontal(_openButton, BoxBuilder.Vertical(dropDownLabel, _dropDown))).
                WithLabel("Desktop Files").
                Build()

    End Function

    Private Shared Function GetFilesOnDesktop() As IEnumerable(Of FileInfo)
        Dim results As ICollection(Of FileInfo) = New Collection(Of FileInfo)

        For Each path As String In Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
            results.Add(New FileInfo(path))
        Next

        Return results
    End Function

    Private Shared Function ConvertFileToDropDownItem(file As FIleInfo, callback As FromControl(Of IPictureDisp)) As Item
        Return New ItemBuilder().
            WithLabel(file.Name).
            WithSuperTip(file.FullName).
            WithImage(Icon.ExtractAssociatedIcon(file.FullName), callback).
            Build()
    End Function

    Private Shared Sub OpenFile(path As String)
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
End Class