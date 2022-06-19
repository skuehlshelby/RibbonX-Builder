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

        Public MustOverride ReadOnly Property ImageType As RibbonImageType

        Public MustOverride Property Image As Object

        Public MustOverride Function AsBuiltInImage() As ImageMSO

        Public MustOverride Function AsIPictureDisp() As IPictureDisp

        Public MustOverride Function AsCachedImage() As String

        Public MustOverride Overrides Function ToString() As String

        Public Shared Function Create(imageName As String) As RibbonImage
            Return New CachedImage(imageName)
        End Function

        Public Shared Function Create(imageMso As ImageMSO) As RibbonImage
            Return New BuiltInImage(imageMso)
        End Function

        Public Shared Function Create(image As IPictureDisp) As RibbonImage
            Return New IPictureDispImage(image)
        End Function

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

            Public Overrides Property Image As Object
                Get
                    Return value
                End Get
                Set
                    Throw New InvalidOperationException("Built-in images (imageMsos) may only be specified at design time.")
                End Set
            End Property

            Public Overrides Function AsBuiltInImage() As ImageMSO
                Return value
            End Function

            Public Overrides Function AsIPictureDisp() As IPictureDisp
                Throw New InvalidOperationException("Built-in images (imageMsos) are not convertible to the type 'IPictureDisp'.")
            End Function

            Public Overrides Function AsCachedImage() As String
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

            Public Overrides Function AsCachedImage() As String
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

            Private ReadOnly imageName As String

            Public Sub New(imageName As String)
                Me.imageName = imageName
            End Sub

            Public Overrides ReadOnly Property ImageType As RibbonImageType
                Get
                    Return RibbonImageType.Cached
                End Get
            End Property

            Public Overrides Property Image As Object
                Get
                    Return imageName
                End Get
                Set
                    Throw New InvalidOperationException("The names of cached images may only be specified at design time. If you control the cache, consider updating the image there.")
                End Set
            End Property

            Public Overrides Function AsBuiltInImage() As ImageMSO
                Throw New InvalidOperationException("Cached images are not convertible to the type 'ImageMSO'.")
            End Function

            Public Overrides Function AsIPictureDisp() As IPictureDisp
                Throw New InvalidOperationException("Cached images are not convertible to the type 'IPictureDisp'.")
            End Function

            Public Overrides Function AsCachedImage() As String
                Return imageName
            End Function

            Public Overrides Function ToString() As String
                Return imageName
            End Function

            Public Overrides Function Equals(obj As Object) As Boolean
                Return imageName.Equals(obj)
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return imageName.GetHashCode()
            End Function

        End Class

    End Class

End Namespace