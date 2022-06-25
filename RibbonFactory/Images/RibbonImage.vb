Imports System.Drawing
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Images.BuiltIn

Namespace Images
    Public MustInherit Class RibbonImage

        Private Sub New()
        End Sub

        Public Enum RibbonImageType As Byte
            BuiltIn
            IPictureDisp
            Cached
        End Enum

        Public Interface ICachedImage

            ReadOnly Property ImageId As String

            ReadOnly Property Image As IPictureDisp

        End Interface

        Public MustOverride ReadOnly Property ImageType As RibbonImageType

        Public MustOverride Property Image As Object

        Public MustOverride Function AsBuiltInImage() As ImageMSO

        Public MustOverride Function AsIPictureDisp() As IPictureDisp

        Public MustOverride Function AsCachedImage() As ICachedImage

        Public MustOverride Overrides Function ToString() As String

        Public Shared Function Create(imageMso As ImageMSO) As RibbonImage
            Return New BuiltInImage(imageMso)
        End Function

        ''' <summary>
        ''' <para>
        '''  Create a RibbonImage with a custom image whose contents will be queried by the MS Office <br/>
        '''  host only once. This retrieval will happen during the 'loadImage' callback. Subsequent control<br/>
        '''  refresh events will not update the image.
        ''' </para>
        ''' <br/>
        ''' <para>
        '''  Use this method when you want to use a custom image that you do not plan on changing at <br/>
        ''' runtime.
        ''' </para>
        ''' </summary>
        ''' <param name="imageId">A unique identifier for this image.</param>
        ''' <param name="bitmap">The image.</param>
        ''' <returns></returns>
        Public Shared Function Create(imageId As String, bitmap As Bitmap) As RibbonImage
            Return New CachedImage(imageId, New AdapterForIPicture(bitmap))
        End Function

        ''' <summary>
        ''' <para>
        '''  Create a RibbonImage with a custom image whose contents will be queried by the MS Office <br/>
        '''  host only once. This retrieval will happen during the 'loadImage' callback. Subsequent control<br/>
        '''  refresh events will not update the image.
        ''' </para>
        ''' <br/>
        ''' <para>
        '''  Use this method when you want to use a custom image that you do not plan on changing at <br/>
        ''' runtime.
        ''' </para>
        ''' </summary>
        ''' <param name="imageId">A unique identifier for this image.</param>
        ''' <param name="icon">The image.</param>
        ''' <returns></returns>
        Public Shared Function Create(imageId As String, icon As Icon) As RibbonImage
            Return New CachedImage(imageId, New AdapterForIPicture(icon))
        End Function

        Public Shared Function Create(image As IPictureDisp) As RibbonImage
            Return New IPictureDispImage(image)
        End Function

        Public Shared Function Create(bitmap As Bitmap) As RibbonImage
            Return Create(New AdapterForIPicture(bitmap))
        End Function

        Public Shared Function Create(icon As Icon) As RibbonImage
            Return Create(New AdapterForIPicture(icon))
        End Function

        Private Class BuiltInImage
            Inherits RibbonImage

            Private value As ImageMSO

            Public Sub New(value As ImageMSO)
                Me.value = value
            End Sub

            Public Overrides ReadOnly Property ImageType As RibbonImageType
                Get
                    Return RibbonImageType.BuiltIn
                End Get
            End Property

            Public Overrides Property Image As Object
                Get
                    Return value
                End Get
                Set
                    If TypeOf Value Is String Then
                        Me.value = ImageMSO.Parse(Of ImageMSO)(Value.ToString())
                    End If
                End Set
            End Property

            Public Overrides Function AsBuiltInImage() As ImageMSO
                Return value
            End Function

            Public Overrides Function AsIPictureDisp() As IPictureDisp
                Throw New InvalidOperationException("Built-in images (imageMsos) are not convertible to the type 'IPictureDisp'.")
            End Function

            Public Overrides Function AsCachedImage() As ICachedImage
                Throw New InvalidOperationException("Built-in images (imageMsos) are not the same as cached images.")
            End Function

            Public Overrides Function ToString() As String
                Return value.ToString()
            End Function

            Public Overrides Function Equals(obj As Object) As Boolean
                Return value.Equals(obj)
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return value.GetHashCode()
            End Function

        End Class

        Private Class IPictureDispImage
            Inherits RibbonImage

            Private value As IPictureDisp

            Public Sub New(value As IPictureDisp)
                Me.value = value
            End Sub

            Public Overrides ReadOnly Property ImageType As RibbonImageType
                Get
                    Return RibbonImageType.IPictureDisp
                End Get
            End Property

            Public Overrides Property Image As Object
                Get
                    Return value
                End Get
                Set
                    If Value IsNot Nothing AndAlso TypeOf Value Is IPictureDisp Then
                        Me.value = DirectCast(Value, IPictureDisp)
                    End If
                End Set
            End Property

            Public Overrides Function AsBuiltInImage() As ImageMSO
                Throw New InvalidOperationException("IPictureDisp images are not convertible to the type 'ImageMSO'.")
            End Function

            Public Overrides Function AsIPictureDisp() As IPictureDisp
                Return value
            End Function

            Public Overrides Function AsCachedImage() As ICachedImage
                Throw New InvalidOperationException("IPictureDisp images are not convertible to the type 'string'.")
            End Function

            Public Overrides Function ToString() As String
                Return String.Empty
            End Function

            Public Overrides Function Equals(obj As Object) As Boolean
                Return value.Equals(obj)
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return value.GetHashCode()
            End Function

        End Class

        Private Class CachedImage
            Inherits RibbonImage
            Implements ICachedImage

            Public Sub New(imageId As String, picture As IPictureDisp)
                Me.ImageId = imageId
                Me.Picture = picture
            End Sub

            Public Overrides ReadOnly Property ImageType As RibbonImageType
                Get
                    Return RibbonImageType.Cached
                End Get
            End Property

            Public Overrides Property Image As Object
                Get
                    Return Picture
                End Get
                Set
                    Throw New InvalidOperationException("The names of cached images may only be specified at design time. If you control the cache, consider updating the image there.")
                End Set
            End Property

            Public ReadOnly Property ImageId As String Implements ICachedImage.ImageId

            Public ReadOnly Property Picture As IPictureDisp Implements ICachedImage.Image

            Public Overrides Function AsBuiltInImage() As ImageMSO
                Throw New InvalidOperationException("Cached images are not convertible to the type 'ImageMSO'.")
            End Function

            Public Overrides Function AsIPictureDisp() As IPictureDisp
                Throw New InvalidOperationException("Cached images are not convertible to the type 'IPictureDisp'.")
            End Function

            Public Overrides Function AsCachedImage() As ICachedImage
                Return Me
            End Function

            Public Overrides Function ToString() As String
                Return ImageId
            End Function

            Public Overrides Function Equals(obj As Object) As Boolean
                Return ImageId.Equals(obj)
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return ImageId.GetHashCode()
            End Function

        End Class

    End Class

End Namespace