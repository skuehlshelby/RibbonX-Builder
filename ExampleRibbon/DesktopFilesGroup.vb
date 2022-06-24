Imports System.Drawing
Imports System.IO
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Images.BuiltIn

Public Class DesktopFilesGroup

    Private ReadOnly _dropDown As DropDown
    Private ReadOnly _openButton As Button
    Private ReadOnly _fileWatcher As FileSystemWatcher

    Public Sub New(ribbon As ICreateCallbacks)
        _fileWatcher = New FileSystemWatcher(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))

        With _fileWatcher
            .EnableRaisingEvents = True
            AddHandler .Created, AddressOf OnFileChanged
            AddHandler .Renamed, AddressOf OnFileChanged
            AddHandler .Deleted, AddressOf OnFileChanged
        End With

        _openButton = New Button(config:=Sub(bb) bb.
            Large().
            WithLabel("Open File").
            WithSuperTip("Open or launch the selected file/program.").
            WithImage(Common.FileOpen).
            OnClick(Sub() OpenFile(DirectCast(_dropDown.Selected.Tag, FileSystemInfo).FullName)).
            RouteClickTo(AddressOf ribbon.OnAction))

        _dropDown = New DropDown(Sub(ddb) ddb.
            WithScreenTip("Desktop Files").
            WithSuperTip("The files on your desktop.").
            AsWideAs("A DropDown This Big").
            GetItemCountFrom(AddressOf ribbon.GetItemCount).
            GetItemIdFrom(AddressOf ribbon.GetItemID).
            GetItemImageFrom(AddressOf ribbon.GetItemImage).
            GetItemLabelFrom(AddressOf ribbon.GetItemLabel).
            GetItemSuperTipFrom(AddressOf ribbon.GetItemSuperTip).
            GetItemScreenTipFrom(AddressOf ribbon.GetItemScreenTip).
            GetSelectedItemIndexFrom(AddressOf ribbon.GetSelectedItemIndex, AddressOf ribbon.OnSelectionChange),
                                 buttons:={_openButton.Clone()})

        Using _dropDown.RefreshSuspension()
            For Each file As FileInfo In GetFilesOnDesktop()
                _dropDown.Add(ConvertFileToDropDownItem(file))
            Next
        End Using
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
                            Using updateBlock As IDisposable = .RefreshSuspension()
                                .Label = e.Name
                                .ScreenTip = e.Name
                                .SuperTip = e.FullPath
                            End Using
                        End With
                        Exit For
                    End If
                Next
        End Select
    End Sub

    Public Function AsGroup() As Group
        Dim label As LabelControl = New LabelControl(config:=Sub(lcb) lcb.WithLabel("        Desktop Files:"))

        Return New Group(Sub(gb) gb.WithLabel("Desktop Files"), {_openButton, Separator.CreateInvisible(), Box.Vertical(label, _dropDown)})
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
        Return New Item(config:=Sub(ib) ib.
            WithLabel(file.Name).
            WithSuperTip(file.FullName).
            WithImage(Icon.ExtractAssociatedIcon(file.FullName)), tag:=file)
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