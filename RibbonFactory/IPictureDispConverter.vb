Imports System.Drawing
Imports System.Runtime.InteropServices

''' <summary>
''' This module supports the creation of custom images for the ribbon.
''' It provides a utility method for converting bitmaps to IPictureDisp.
''' </summary>
Public Module PictureDispConverter

    <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", CharSet:=CharSet.Ansi, ExactSpelling:=True, PreserveSig:=True)>
    Private Function OleCreatePictureIndirect(ByRef pictdesc As Pictdesc, ByRef riid As Guid, fOwn As Boolean, <MarshalAs(UnmanagedType.IDispatch)> ByRef lplpvObj As Object) As Integer
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure Pictdesc
        Public SizeOfStruct As Integer
        Public PicType As Integer
        Public bitmapHandle As IntPtr
        Public Hpal As IntPtr
        Public Padding As Integer
        Public Sub New(hBmp As IntPtr)
            SizeOfStruct = Marshal.SizeOf(GetType(Pictdesc))
            PicType = 1
            bitmapHandle = hBmp
            Hpal = IntPtr.Zero
            Padding = 0
        End Sub
    End Structure

    Public Function ImageToPictureDisp(input As Bitmap) As stdole.IPictureDisp
        Dim pictDesc As Pictdesc = New Pictdesc(input.GetHbitmap())
        Dim pic As Object = Nothing
        OleCreatePictureIndirect(pictDesc, GetType(stdole.IPictureDisp).GUID, True, pic)
        Return CType(pic, stdole.IPictureDisp)
    End Function
End Module
