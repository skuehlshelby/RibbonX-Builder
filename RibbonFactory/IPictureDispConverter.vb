Imports System.Runtime.InteropServices

Public Class IPictureDispConverter

    Private Shared IPictureDispGUID As Guid = GetType(stdole.IPictureDisp).GUID

    <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)>
    Private Shared Function OleCreatePictureIndirect(<MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, ByRef iid As Guid, ByVal fOwn As Boolean)
    End Function

    Public Shared Function Convert(ByVal Icon As Drawing.Icon) As stdole.IPictureDisp
        Dim pictIcon As PICTDESC.Icon = New PICTDESC.Icon(Icon)
        Return IPictureDispConverter.OleCreatePictureIndirect(pictIcon, IPictureDispGUID, True)
    End Function

    Public Shared Function Convert(ByVal Image As Drawing.Image) As stdole.IPictureDisp
        Dim Bitmap As Drawing.Bitmap = If(TypeOf Image Is Drawing.Bitmap, DirectCast(Image, Drawing.Bitmap), New Drawing.Bitmap(Image))
        Dim pictBit As PICTDESC.Bitmap = New PICTDESC.Bitmap(Bitmap)
        Return IPictureDispConverter.OleCreatePictureIndirect(pictBit, IPictureDispGUID, True)
    End Function

    Private Class PICTDESC
        Public Enum PICTURE_TYPE As SByte
            UNINITIALIZED = -1
            NONE = 0
            BITMAP = 1
            METAFILE = 2
            ICON = 3
            ENHMETAFILE = 4
        End Enum

        <StructLayout(LayoutKind.Sequential)>
        Public Class Icon
            Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
            Friend PictureType As Integer = PICTURE_TYPE.ICON
            Friend IconHandle As IntPtr = IntPtr.Zero
            Friend unused1 As Integer = 0
            Friend unused2 As Integer = 0

            Friend Sub New(ByVal Icon As Drawing.Icon)
                IconHandle = Icon.ToBitmap().GetHicon()
            End Sub
        End Class

        <StructLayout(LayoutKind.Sequential)>
        Public Class Bitmap
            Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
            Friend PictureType As Integer = PICTURE_TYPE.BITMAP
            Friend BitmapHandle As IntPtr = IntPtr.Zero
            Friend PaletteHandle As IntPtr = IntPtr.Zero
            Friend unused As Integer = 0

            Friend Sub New(ByVal Bitmap As Drawing.Bitmap)
                BitmapHandle = Bitmap.GetHbitmap()
            End Sub
        End Class

    End Class
End Class