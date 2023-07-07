Imports System.Drawing
Imports System.IO
Imports RibbonX
Imports RibbonX.Callbacks
Imports RibbonX.Images.BuiltIn

Public Class DesktopFilesGroup

    Private ReadOnly _dropDown As IDropDown
    Private ReadOnly _openButton As IButton
    Private ReadOnly _fileWatcher As FileSystemWatcher

    Public Sub New(ribbon As ICreateCallbacks)
        _fileWatcher = New FileSystemWatcher(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))

        With _fileWatcher
            .EnableRaisingEvents = True
            AddHandler .Created, AddressOf OnFileChanged
            AddHandler .Renamed, AddressOf OnFileChanged
            AddHandler .Deleted, AddressOf OnFileChanged
        End With

        _openButton = RxApi.Button(Sub(bb) bb.
            Large().
            WithLabel("Open File").
            WithSuperTip("Open or launch the selected file/program.").
            WithImage(Common.FileOpen).
            OnClick(AddressOf ribbon.OnAction, Sub() OpenFile(DirectCast(_dropDown.Selected.Tag, FileSystemInfo).FullName)))

        _dropDown = RxApi.DropDown(Sub(ddb) ddb.
            WithScreenTip("Desktop Files").
            WithSuperTip("The files on your desktop.").
            AsWideAs("A DropDown This Big").
            GetItemCountFrom(AddressOf ribbon.GetItemCount).
            GetItemIdFrom(AddressOf ribbon.GetItemID).
            GetItemImageFrom(AddressOf ribbon.GetItemImage).
            GetItemLabelFrom(AddressOf ribbon.GetItemLabel).
            GetItemSuperTipFrom(AddressOf ribbon.GetItemSuperTip).
            GetItemScreenTipFrom(AddressOf ribbon.GetItemScreenTip).
            WithButtons(DirectCast(_openButton.Clone(), IButton)).
            GetSelectedItemIndexFrom(AddressOf ribbon.GetSelectedItemIndex, AddressOf ribbon.OnSelectionChange))

        Using _dropDown.SuspendRefreshing()
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
                For Each item As IItem In _dropDown
                    If Not DirectCast(item.Tag, FileInfo).Exists Then
                        _dropDown.Remove(item)
                        Exit For
                    End If
                Next
            Case WatcherChangeTypes.Renamed
                For Each item As IItem In _dropDown
                    If Not DirectCast(item.Tag, FileInfo).Exists Then
                        With item
                            Using updateBlock As IDisposable = .SuspendRefreshing()
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

    Public Function AsGroup() As IGroup
        Dim label As ILabelControl = RxApi.LabelControl(Sub(lcb) lcb.WithLabel("        Desktop Files:"))

        Return RxApi.Group(Sub(gb) gb.WithLabel("Desktop Files").WithControls(_openButton, RxApi.Separator(Sub(b) b.Invisible()), RxApi.Box(Sub(b) b.Vertical.WithControls(_dropDown))))
    End Function

    Private Shared Function GetFilesOnDesktop() As IEnumerable(Of FileInfo)
        Return Directory.
            GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)).
            Select(Function(filePath) New FileInfo(filePath))
    End Function

    Private Shared Function ConvertFileToDropDownItem(filePath As String) As IItem
        Return ConvertFileToDropDownItem(New FileInfo(filePath))
    End Function

    Private Shared Function ConvertFileToDropDownItem(file As FileSystemInfo) As IItem
        Return RxApi.Item(Sub(ib) ib.
            WithLabel(file.Name).
            WithSuperTip(file.FullName).
            WithImage(Icon.ExtractAssociatedIcon(file.FullName)).
            WithTag(file))
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