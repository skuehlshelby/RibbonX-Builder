Imports System.Drawing
Imports System.Runtime.InteropServices

'This module supports the creation of custom images for the ribbon.
Public Module IPictureDispConverter

    <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", CharSet:=CharSet.Ansi, ExactSpelling:=True, PreserveSig:=True)>
    Private Function OleCreatePictureIndirect(ByRef pictdesc As PICTDESC, ByRef riid As Guid, fOwn As Boolean, <MarshalAs(UnmanagedType.IDispatch)> ByRef lplpvObj As Object) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure PICTDESC
        Public SizeOfStruct As Integer
        Public PicType As Integer
        Public Hbitmap As IntPtr
        Public Hpal As IntPtr
        Public Padding As Integer
        Public Sub New(hBmp As IntPtr)
            SizeOfStruct = Marshal.SizeOf(GetType(PICTDESC))
            PicType = 1
            Hbitmap = hBmp
            Hpal = IntPtr.Zero
            Padding = 0
        End Sub
    End Structure

    Public Function ImageToPictureDisp(Input As Bitmap) As stdole.IPictureDisp
        Dim PictDesc As PICTDESC = New PICTDESC(Input.GetHbitmap())
        Dim Pic As Object = Nothing
        OleCreatePictureIndirect(PictDesc, GetType(stdole.IPictureDisp).GUID, True, Pic)
        Return CType(Pic, stdole.IPictureDisp)
    End Function
End Module
