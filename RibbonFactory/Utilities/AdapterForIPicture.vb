Imports System.Drawing
Imports System.Runtime.InteropServices
Imports stdole

Namespace Utilities

    Public Class AdapterForIPicture
        Implements IPicture
        Implements IPictureDisp
        Implements IDisposable

        Private _image As Bitmap
        Private _handle As IntPtr
        Private _disposedValue As Boolean

        <DllImport("gdi32.dll")>
        Private Shared Sub DeleteObject(handle As IntPtr)
        End Sub

        Public Sub New(image As Bitmap)
            _image = image
        End Sub

        Public Sub New(image As Icon)
            _image = image.ToBitmap()
        End Sub

        Public Sub Render(hdc As Integer, x As Integer, y As Integer, cx As Integer, cy As Integer, xSrc As Integer, ySrc As Integer, cxSrc As Integer, cySrc As Integer, prcWBounds As IntPtr) Implements IPictureDisp.Render, IPicture.Render
            Using graphics As Graphics = Graphics.FromHdc(New IntPtr(hdc))
                graphics.DrawImage(_image, new Rectangle(x, y, cx, cy), xSrc, ySrc, cxSrc, cySrc, GraphicsUnit.Pixel)
            End Using
        End Sub

        Public ReadOnly Property Handle As Integer Implements IPictureDisp.Handle, IPicture.Handle
            Get 
                If _handle = IntPtr.Zero Then
                    _handle = _image.GetHbitmap()
                End If

                Return _handle.ToInt32()
            End Get
        End Property

        Public Property hPal As Integer Implements IPictureDisp.hPal, IPicture.hPal
            Get
                Return 0
            End Get
            Set
                Return
            End Set
        End Property

        Public ReadOnly Property Type As Short Implements IPictureDisp.Type, IPicture.Type
            Get
                Return 1 'Numeric code for bitmap
            End Get
        End Property

        Public ReadOnly Property Width As Integer Implements IPictureDisp.Width, IPicture.Width
            Get
                Return _image.Width
            End Get
        End Property

        Public ReadOnly Property Height As Integer Implements IPictureDisp.Height, IPicture.Height
            Get
                Return _image.Height
            End Get
        End Property

        Public Sub SelectPicture(hdcIn As Integer, ByRef phdcOut As Integer, ByRef phbmpOut As Integer) Implements IPicture.SelectPicture
            phdcOut = 0
            phbmpOut = 0
        End Sub

        Public Sub PictureChanged() Implements IPicture.PictureChanged
            Return
        End Sub

        Public Sub SaveAsFile(pstm As IntPtr, fSaveMemCopy As Boolean, ByRef pcbSize As Integer) Implements IPicture.SaveAsFile
            pcbSize = 0
        End Sub

        Public Sub SetHdc(hdc As Integer) Implements IPicture.SetHdc
            Return
        End Sub

        Public ReadOnly Property CurDC As Integer Implements IPicture.CurDC
            Get
                Return 0
            End Get
        End Property

        Public Property KeepOriginalFormat As Boolean Implements IPicture.KeepOriginalFormat
            Get
                Return False
            End Get
            Set
                Return
            End Set
        End Property

        Public ReadOnly Property Attributes As Integer Implements IPicture.Attributes
            Get
                Return 0
            End Get
        End Property

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposedValue Then
                If disposing Then
                    _image.Dispose()
                End If

                If _handle <> IntPtr.Zero Then
                    DeleteObject(_handle)
                End If
            
                _image = Nothing
                _handle = Nothing
                _disposedValue = True
            End If
        End Sub

        Protected Overrides Sub Finalize()
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=False)
            MyBase.Finalize()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End NameSpace