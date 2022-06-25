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

        Public MustOverride ReadOnly Property Image As Object

        Public MustOverride Function AsBuiltInImage() As ImageMSO

        Public MustOverride Function AsIPictureDisp() As IPictureDisp

        Public MustOverride Function AsCachedImage() As ICachedImage

        Public MustOverride Overrides Function ToString() As String

        Public Shared Function Create(imageMso As ImageMSO) As RibbonImage
            Return New BuiltInImage(imageMso)
        End Function

        Public Shared Function Create(imageId As String, bitmap As Bitmap) As RibbonImage
            Return New CachedImage(imageId, New AdapterForIPicture(bitmap))
        End Function

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

        Public Shared Widening Operator CType(image As ImageMSO) As RibbonImage
            Return Create(image)
        End Operator

        Public Shared Widening Operator CType(image As Bitmap) As RibbonImage
            Return Create(image)
        End Operator

        Public Shared Widening Operator CType(image As Icon) As RibbonImage
            Return Create(image)
        End Operator

        Private Class BuiltInImage
            Inherits RibbonImage

            Private ReadOnly value As ImageMSO

            Public Sub New(value As ImageMSO)
                Me.value = value
            End Sub

            Public Overrides ReadOnly Property ImageType As RibbonImageType
                Get
                    Return RibbonImageType.BuiltIn
                End Get
            End Property

            Public Overrides ReadOnly Property Image As Object
                Get
                    Return value
                End Get
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

            Private ReadOnly value As IPictureDisp

            Public Sub New(value As IPictureDisp)
                Me.value = value
            End Sub

            Public Overrides ReadOnly Property ImageType As RibbonImageType
                Get
                    Return RibbonImageType.IPictureDisp
                End Get
            End Property

            Public Overrides ReadOnly Property Image As Object
                Get
                    Return value
                End Get
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

            Public Overrides ReadOnly Property Image As Object
                Get
                    Return Picture
                End Get
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