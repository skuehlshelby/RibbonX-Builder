Imports System.Drawing
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Controls.EventArgs
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.SimpleTypes
Imports RibbonX.Callbacks
Imports RibbonX.Builders.Internal.Base
Imports RibbonX.Images.BuiltIn

Namespace Builders

    Friend NotInheritable Class EditBoxBuilder
        Inherits BuilderBase(Of EditBox)
        Implements IEditBoxBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function Enabled() As IEditBoxBuilder Implements IEditBoxBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As IEditBoxBuilder Implements IEditBoxBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As IEditBoxBuilder Implements IEditBoxBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As IEditBoxBuilder Implements IEditBoxBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As IEditBoxBuilder Implements IEditBoxBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As IEditBoxBuilder Implements IEditBoxBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As IEditBoxBuilder Implements IEditBoxBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As IEditBoxBuilder Implements IEditBoxBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IEditBoxBuilder Implements IEditBoxBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IEditBoxBuilder Implements IEditBoxBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IEditBoxBuilder Implements IEditBoxBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IEditBoxBuilder Implements IEditBoxBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IEditBoxBuilder Implements IEditBoxBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IEditBoxBuilder Implements IEditBoxBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As IEditBoxBuilder Implements IEditBoxBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IEditBoxBuilder Implements IEditBoxBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO) As IEditBoxBuilder Implements IEditBoxBuilder.WithImage
            ImageBase(image)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As IEditBoxBuilder Implements IEditBoxBuilder.WithImage
            Throw New NotImplementedException()
        End Function

        Private Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IEditBoxBuilder Implements IEditBoxBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IEditBoxBuilder Implements IEditBoxBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Bitmap) As IEditBoxBuilder Implements IEditBoxBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Icon) As IEditBoxBuilder Implements IEditBoxBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Private Function ShowLabel() As IEditBoxBuilder Implements IEditBoxBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As IEditBoxBuilder Implements IEditBoxBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabel() As IEditBoxBuilder Implements IEditBoxBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As IEditBoxBuilder Implements IEditBoxBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Private Function WithMaximumInputLength(maxLength As Byte) As IEditBoxBuilder Implements IEditBoxBuilder.WithMaximumInputLength
            MaxLengthBase(maxLength)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As IEditBoxBuilder Implements IEditBoxBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As IEditBoxBuilder Implements IEditBoxBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As IEditBoxBuilder Implements IEditBoxBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As IEditBoxBuilder Implements IEditBoxBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Private Function AsWideAs(sizeString As String) As IEditBoxBuilder Implements IEditBoxBuilder.AsWideAs
            SizeStringBase(sizeString)
            Return Me
        End Function

        Private Function WithId(id As String) As IEditBoxBuilder Implements IEditBoxBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As IEditBoxBuilder Implements IEditBoxBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As IEditBoxBuilder Implements IEditBoxBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function WithText(text As String, getText As FromControl(Of String), setText As TextChanged) As IEditBoxBuilder Implements IEditBoxBuilder.WithText
            TextBase(text, getText, setText)
            Return Me
        End Function

        Private Function BeforeTextChange(ParamArray actions() As EventHandler(Of BeforeTextChangeEventArgs)) As IEditBoxBuilder Implements IEditBoxBuilder.BeforeTextChange
            AddBeforeTextChangeHandlers(actions)
            Return Me
        End Function

        Private Function AfterTextChange(ParamArray actions() As EventHandler(Of TextChangeEventArgs)) As IEditBoxBuilder Implements IEditBoxBuilder.AfterTextChange
            AddOnTextChangeHandlers(actions)
            Return Me
        End Function

    End Class

End Namespace