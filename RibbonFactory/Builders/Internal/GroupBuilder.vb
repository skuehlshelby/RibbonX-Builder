Imports System.Drawing
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.SimpleTypes
Imports RibbonX.Callbacks
Imports RibbonX.Builders.Internal.Base
Imports RibbonX.Images.BuiltIn

Namespace Builders
    Friend NotInheritable Class GroupBuilder
        Inherits BuilderBase(Of Group)
        Implements IGroupBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Public Function WithId(id As String) As IGroupBuilder Implements IGroupBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As IGroupBuilder Implements IGroupBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As IGroupBuilder Implements IGroupBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Public Function Visible() As IGroupBuilder Implements IGroupBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IGroupBuilder Implements IGroupBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As IGroupBuilder Implements IGroupBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IGroupBuilder Implements IGroupBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As IGroupBuilder Implements IGroupBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IGroupBuilder Implements IGroupBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Bitmap) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Icon) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As IGroupBuilder Implements IGroupBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As IGroupBuilder Implements IGroupBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As IGroupBuilder Implements IGroupBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As IGroupBuilder Implements IGroupBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IGroupBuilder Implements IGroupBuilder.WithLabel
            LabelBase(label)

            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IGroupBuilder Implements IGroupBuilder.WithLabel
            LabelBase(label, callback)

            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As IGroupBuilder Implements IGroupBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IGroupBuilder Implements IGroupBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As IGroupBuilder Implements IGroupBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IGroupBuilder Implements IGroupBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

    End Class

End Namespace