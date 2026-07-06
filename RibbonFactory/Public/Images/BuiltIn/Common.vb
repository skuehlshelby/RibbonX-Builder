Namespace Images.BuiltIn
    Public NotInheritable Class Common
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property AttachFile As Common = New Common(1, NameOf(AttachFile))

        Public Shared ReadOnly Property Clear As Common = New Common(2, NameOf(Clear))

        Public Shared ReadOnly Property Cut As Common = New Common(3, NameOf(Cut))

        Public Shared ReadOnly Property Delete As Common = New Common(4, NameOf(Delete))

        Public Shared ReadOnly Property DollarSign As Common = New Common(5, NameOf(DollarSign))

        Public Shared ReadOnly Property FileFind As Common = New Common(6, NameOf(FileFind))

        Public Shared ReadOnly Property FileNew As Common = New Common(7, NameOf(FileNew))

        Public Shared ReadOnly Property FileOpen As Common = New Common(8, NameOf(FileOpen))

        Public Shared ReadOnly Property FileSave As Common = New Common(9, NameOf(FileSave))

        Public Shared ReadOnly Property FileSaveAs As Common = New Common(10, NameOf(FileSaveAs))

        Public Shared ReadOnly Property FillDown As Common = New Common(11, NameOf(FillDown))

        Public Shared ReadOnly Property Folder As Common = New Common(12, NameOf(Folder))

        Public Shared ReadOnly Property HappyFace As Common = New Common(13, NameOf(HappyFace))

        Public Shared ReadOnly Property Paste As Common = New Common(14, NameOf(Paste))

        Public Shared ReadOnly Property Refresh As Common = New Common(15, NameOf(Refresh))

        Public Shared ReadOnly Property Repeat As Common = New Common(16, NameOf(Repeat))

        Public Shared ReadOnly Property SadFace As Common = New Common(17, NameOf(SadFace))

        Public Shared ReadOnly Property SaveAll As Common = New Common(18, NameOf(SaveAll))

        Public Shared ReadOnly Property TraceError As Common = New Common(19, NameOf(TraceError))

        Public Shared ReadOnly Property Undo As Common = New Common(20, NameOf(Undo))

        Public Overrides Function Clone() As Object
            Return New Common(value, name)
        End Function

    End Class

End Namespace
