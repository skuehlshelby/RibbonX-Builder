Namespace Enums.ImageMSO
    Public NotInheritable Class Common
        Inherits ImageMSO
        
        Private Sub New(name As String)
            MyBase.New(name)
        End Sub
        
        Public Shared ReadOnly Property AttachFile As Common = New Common(NameOf(AttachFile))

        Public Shared ReadOnly Property Clear As Common = New Common(NameOf(Clear))

        Public Shared ReadOnly Property Cut As Common = New Common(NameOf(Cut))

        Public Shared ReadOnly Property Delete As Common = New Common(NameOf(Delete))

        Public Shared ReadOnly Property DollarSign As Common = New Common(NameOf(DollarSign))

        Public Shared ReadOnly Property FileFind As Common = New Common(NameOf(FileFind))

        Public Shared ReadOnly Property FileNew As Common = New Common(NameOf(FileNew))

        Public Shared ReadOnly Property FileOpen As Common = New Common(NameOf(FileOpen))

        Public Shared ReadOnly Property FileSave As Common = New Common(NameOf(FileSave))

        Public Shared ReadOnly Property FileSaveAs As Common = New Common(NameOf(FileSaveAs))

        Public Shared ReadOnly Property FillDown As Common = New Common(NameOf(FillDown))

        Public Shared ReadOnly Property Folder As Common = New Common(NameOf(Folder))

        Public Shared ReadOnly Property HappyFace As Common = New Common(NameOf(HappyFace))

        Public Shared ReadOnly Property Paste As Common = New Common(NameOf(Paste))

        Public Shared ReadOnly Property Refresh As Common = New Common(NameOf(Refresh))

        Public Shared ReadOnly Property Repeat As Common = New Common(NameOf(Repeat))

        Public Shared ReadOnly Property SadFace As Common = New Common(NameOf(SadFace))

        Public Shared ReadOnly Property SaveAll As Common = New Common(NameOf(SaveAll))

        Public Shared ReadOnly Property TraceError As Common = New Common(NameOf(TraceError))

        Public Shared ReadOnly Property Undo As Common = New Common(NameOf(Undo))
        
    End Class
End NameSpace