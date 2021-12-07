Imports System.Drawing
Imports System.Runtime.InteropServices

''' <summary>
''' This module supports the creation of custom images for the ribbon.
''' It provides a utility method for converting bitmaps to IPictureDisp.
''' </summary>
Public Module ImageConverter
    <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
    Private Function OleCreatePictureIndirect(<MarshalAs(UnmanagedType.AsAny)>picdesc As Object, ByRef iid As Guid, <MarshalAs(UnmanagedType.Bool)>fOwn As Boolean) As stdole.IPictureDisp
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

    Public Function ConvertToIPictureDisplay(input As Bitmap) As stdole.IPictureDisp
        Dim pictDesc As Pictdesc = New Pictdesc(input.GetHbitmap())
        
        Return OleCreatePictureIndirect(pictDesc, GetType(stdole.IPictureDisp).GUID, True)
    End Function
End Module