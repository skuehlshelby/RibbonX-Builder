Namespace Enums.ImageMSO
    Public NotInheritable Class Common
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property AttachFile As Common = New Common(1, NameOf(AttachFile))

        Public Shared ReadOnly Property Clear As Common = New Common(1, NameOf(Clear))

        Public Shared ReadOnly Property Cut As Common = New Common(1, NameOf(Cut))

        Public Shared ReadOnly Property Delete As Common = New Common(1, NameOf(Delete))

        Public Shared ReadOnly Property DollarSign As Common = New Common(1, NameOf(DollarSign))

        Public Shared ReadOnly Property FileFind As Common = New Common(1, NameOf(FileFind))

        Public Shared ReadOnly Property FileNew As Common = New Common(1, NameOf(FileNew))

        Public Shared ReadOnly Property FileOpen As Common = New Common(1, NameOf(FileOpen))

        Public Shared ReadOnly Property FileSave As Common = New Common(1, NameOf(FileSave))

        Public Shared ReadOnly Property FileSaveAs As Common = New Common(1, NameOf(FileSaveAs))

        Public Shared ReadOnly Property FillDown As Common = New Common(1, NameOf(FillDown))

        Public Shared ReadOnly Property Folder As Common = New Common(1, NameOf(Folder))

        Public Shared ReadOnly Property HappyFace As Common = New Common(1, NameOf(HappyFace))

        Public Shared ReadOnly Property Paste As Common = New Common(1, NameOf(Paste))

        Public Shared ReadOnly Property Refresh As Common = New Common(1, NameOf(Refresh))

        Public Shared ReadOnly Property Repeat As Common = New Common(1, NameOf(Repeat))

        Public Shared ReadOnly Property SadFace As Common = New Common(1, NameOf(SadFace))

        Public Shared ReadOnly Property SaveAll As Common = New Common(1, NameOf(SaveAll))

        Public Shared ReadOnly Property TraceError As Common = New Common(1, NameOf(TraceError))

        Public Shared ReadOnly Property Undo As Common = New Common(1, NameOf(Undo))

        Public Overrides Function Clone() As Object
            Return New Common(value, name)
        End Function

    End Class

End NameSpace