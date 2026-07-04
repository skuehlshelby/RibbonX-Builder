Imports System.Drawing
Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.InternalApi
Imports RibbonX.SimpleTypes

Namespace Builders
    Friend NotInheritable Class GroupBuilder
        Inherits BuilderBase(Of Group)
        Implements IGroupBuilder

        Private ReadOnly items As ISet(Of IRibbonElement) = New HashSet(Of IRibbonElement)()

        Private Function WithId(id As String) As IGroupBuilder Implements IGroupBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As IGroupBuilder Implements IGroupBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As IGroupBuilder Implements IGroupBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function Visible() As IGroupBuilder Implements IGroupBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As IGroupBuilder Implements IGroupBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As IGroupBuilder Implements IGroupBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As IGroupBuilder Implements IGroupBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As IGroupBuilder Implements IGroupBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IGroupBuilder Implements IGroupBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image)
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(imageId As String, image As Bitmap) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Private Function WithImage(imageId As String, image As Icon) As IGroupBuilder Implements IGroupBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As IGroupBuilder Implements IGroupBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As IRibbonElement) As IGroupBuilder Implements IGroupBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As IGroupBuilder Implements IGroupBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As IRibbonElement) As IGroupBuilder Implements IGroupBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IGroupBuilder Implements IGroupBuilder.WithLabel
            LabelBase(label)

            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IGroupBuilder Implements IGroupBuilder.WithLabel
            LabelBase(label, callback)

            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IGroupBuilder Implements IGroupBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IGroupBuilder Implements IGroupBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IGroupBuilder Implements IGroupBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IGroupBuilder Implements IGroupBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Public Overrides Function Build() As Group
            Return New Group(ControlProperties, items.ToArray(), Tag)
        End Function

        Private Function WithTag(tag As Object) As IGroupBuilder Implements ITag(Of IGroupBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Private Function FromTemplate(template As IRibbonElement) As IGroupBuilder Implements IGroupBuilder.FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Private Function WithControls(ParamArray controls() As IGroupAddable) As IGroupBuilder Implements IGroupBuilder.WithControls
            For Each control As IGroupAddable In controls
                items.Add(control)
            Next
            Return Me
        End Function

        Public Shared Function FromAction(action As Action(Of IGroupBuilder)) As IGroup
            Dim instance As GroupBuilder = New GroupBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function
    End Class

End Namespace