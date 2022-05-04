Imports System.Drawing
Imports System.Reflection

Friend Module Utilities

    Public Sub MessageBox(prompt As String)
        MsgBox(prompt, MsgBoxStyle.OkOnly, My.Application.Info.Title)
    End Sub

    Public Function LoadBitmap(path As String) As Bitmap
        Return New Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream(path))
    End Function

End Module