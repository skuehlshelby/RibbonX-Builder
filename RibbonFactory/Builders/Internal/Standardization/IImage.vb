Imports System.Drawing
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Callbacks
Imports RibbonX.Images.BuiltIn

Namespace Builders.Internal.Standardization

    Public Interface IImage(Of T)

        Function WithImage(image As ImageMSO) As T

        Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As T

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
        ''' <param name="image">The image.</param>
        ''' <returns></returns>
        Function WithImage(imageId As String, image As Bitmap) As T

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
        ''' <param name="image">The image.</param>
        ''' <returns></returns>
        Function WithImage(imageId As String, image As Icon) As T

        Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As T

        Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As T

    End Interface

End Namespace