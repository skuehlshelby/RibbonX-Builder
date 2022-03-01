﻿Imports System.Drawing
Imports System.IO
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls

Public Class DesktopFilesGroup
    
    Private ReadOnly _dropDown As DropDown
    Private ReadOnly _openButton As Button
    Private Readonly _fileWatcher As FileSystemWatcher

    Public Sub New(ribbon As ICreateCallbacks)
        _fileWatcher = New FileSystemWatcher(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
        
        With _fileWatcher
            .EnableRaisingEvents = True
            AddHandler .Created, AddressOf OnFileChanged
            AddHandler .Renamed, AddressOf OnFileChanged
            AddHandler .Deleted, AddressOf OnFileChanged
        End With

        _openButton = New ButtonBuilder().
            Large().
            WithLabel("Open File").
            WithSuperTip("Open or launch the selected file/program.").
            WithImage(Enums.ImageMSO.Common.FileOpen).
            OnClick(Sub() OpenFile(DirectCast(_dropDown.Selected.Tag, FileSystemInfo).FullName)).
            WithClickCallback(AddressOf ribbon.OnAction).
            Build()

        _dropDown = New DropDownBuilder().
            Add(New ButtonBuilder(_openButton).Build()).
            WithScreenTip("Desktop Files").
            WithSuperTip("The files on your desktop.").
            AsWideAs("A DropDown This Big").
            GetItemCountFrom(AddressOf ribbon.GetItemCount).
            GetItemIdFrom(AddressOf ribbon.GetItemID).
            GetItemImageFrom(AddressOf ribbon.GetItemImage).
            GetItemLabelFrom(AddressOf ribbon.GetItemLabel).
            GetItemSuperTipFrom(AddressOf ribbon.GetItemSuperTip).
            GetItemScreenTipFrom(AddressOf ribbon.GetItemScreenTip).
            GetSelectedItemIndexFrom(AddressOf ribbon.GetSelectedItemIndex, AddressOf ribbon.OnSelectionChange).
            Build()

        For Each file As FileInfo In GetFilesOnDesktop()
            _dropDown.Add(ConvertFileToDropDownItem(file))
        Next
    End Sub

    Private Sub OnFileChanged(sender As Object, e As FileSystemEventArgs)
        Select Case e.ChangeType
            Case WatcherChangeTypes.Created
                _dropDown.Add(ConvertFileToDropDownItem(e.FullPath))
            Case WatcherChangeTypes.Deleted
                For Each item As Item In _dropDown
                    If Not DirectCast(item.Tag, FileInfo).Exists Then
                        _dropDown.Remove(item)
                        Exit For
                    End If
                Next
            Case WatcherChangeTypes.Renamed
                For Each item As Item In _dropDown
                    If Not DirectCast(item.Tag, FileInfo).Exists Then
                        With item
                            .SuspendAutomaticRefreshing()
                            .Label = e.Name
                            .ScreenTip = e.Name
                            .SuperTip = e.FullPath
                            .ResumeAutomaticRefreshing()
                        End With
                        Exit For
                    End If
                Next
        End Select
    End Sub

    Public Function AsGroup() As Group
        Dim label As LabelControl = New LabelControlBuilder().
            WithLabel("        Desktop Files:").
            Build()

        Dim separator As Separator = New SeparatorBuilder().
            Invisible().
            Build()

        Return New GroupBuilder().
                WithControls(_openButton, separator, BoxBuilder.Vertical(label, _dropDown)).
                WithLabel("Desktop Files").
                Build()
    End Function

    Private Shared Function GetFilesOnDesktop() As IEnumerable(Of FileInfo)
        Return Directory.
            GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)).
            Select(Function(filePath) New FileInfo(filePath))
    End Function

    Private Shared Function ConvertFileToDropDownItem(filePath As String) As Item
        Return ConvertFileToDropDownItem(New FileInfo(filePath))
    End Function

    Private Shared Function ConvertFileToDropDownItem(file As FileSystemInfo) As Item
        Return New ItemBuilder().
            WithLabel(file.Name).
            WithSuperTip(file.FullName).
            WithImage(Icon.ExtractAssociatedIcon(file.FullName)).
            Build(file)
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