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

    Private NotInheritable Class PICTDESC
        Private Sub New()
        End Sub

        Private Enum PicType As Short
            Uninitialized = -1
            None = 0
            Bitmap = 1
            MetaFile = 2
            Icon = 3
            EnhMetaFile = 4
        End Enum

        <StructLayout(LayoutKind.Sequential)>
        Public Structure Bitmap
            Private ReadOnly SizeOfStruct As Integer
            Private ReadOnly PicType As Integer
            Private ReadOnly BitmapHandle As IntPtr
            Private ReadOnly PaletteHandle As IntPtr
            Private ReadOnly Padding As Integer

            Public Sub New(input As Drawing.Bitmap)
                SizeOfStruct = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                PicType = PICTDESC.PicType.Bitmap
                BitmapHandle = input.GetHbitmap()
                PaletteHandle = IntPtr.Zero
                Padding = 0
            End Sub
        End Structure


        <StructLayout(LayoutKind.Sequential)>
        Public Structure Icon
            Private ReadOnly SizeOfStruct As Integer
            Private ReadOnly PicType As Integer
            Private ReadOnly IconHandle As IntPtr
            Private ReadOnly Padding As Integer
            Private ReadOnly MorePadding As Integer

            Public Sub New(input As Drawing.Icon)
                SizeOfStruct = Marshal.SizeOf(GetType(PICTDESC.Icon))
                picType = PICTDESC.PicType.Icon
                IconHandle = input.ToBitmap().GetHicon()
                Padding = 0
                MorePadding = 0
            End Sub
        End Structure
    End Class

    Public Function ConvertToIPictureDisplay(input As Bitmap) As stdole.IPictureDisp
        Return OleCreatePictureIndirect(New PICTDESC.Bitmap(input), GetType(stdole.IPictureDisp).GUID, True)
    End Function

    Public Function ConvertToIPictureDisplay(input As Icon) As stdole.IPictureDisp
        Return OleCreatePictureIndirect(New PICTDESC.Icon(input), GetType(stdole.IPictureDisp).GUID, True)
    End Function
End Module