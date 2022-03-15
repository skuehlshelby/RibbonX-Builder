Imports System.Drawing
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders
    Friend NotInheritable Class ButtonBuilder
        Implements IButtonBuilder

        Private ReadOnly _builder As ControlBuilder

        Public Sub New()
            _builder = ControlBuilder.Init(Of Button)()
        End Sub

        Public Sub New(template As RibbonElement)
            _builder = ControlBuilder.Init(Of Button)(template)
        End Sub

        Public Function Build() As AttributeSet
            Return _builder.Build()
        End Function

        Public Function WithId(id As String) As IButtonBuilder Implements IButtonBuilder.WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As IButtonBuilder Implements IButtonBuilder.WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As IButtonBuilder Implements IButtonBuilder.WithIdMso
            _builder.WithId(mso)
            Return Me
        End Function

        Public Function Enabled() As IButtonBuilder Implements IButtonBuilder.Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As IButtonBuilder Implements IButtonBuilder.Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As IButtonBuilder Implements IButtonBuilder.Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As IButtonBuilder Implements IButtonBuilder.Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IButtonBuilder Implements IButtonBuilder.WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IButtonBuilder Implements IButtonBuilder.WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As IButtonBuilder Implements IButtonBuilder.WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As IButtonBuilder Implements IButtonBuilder.WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As IButtonBuilder Implements IButtonBuilder.ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As IButtonBuilder Implements IButtonBuilder.HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function Large() As IButtonBuilder Implements IButtonBuilder.Large
            _builder.Large()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As IButtonBuilder Implements IButtonBuilder.Large
            _builder.Large(callback)
            Return Me
        End Function

        Public Function Normal() As IButtonBuilder Implements IButtonBuilder.Normal
            _builder.Normal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As IButtonBuilder Implements IButtonBuilder.Normal
            _builder.Normal(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As IButtonBuilder Implements IButtonBuilder.WithImage
            _builder.WithImage(image)
            _builder.ShowImage()
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IButtonBuilder Implements IButtonBuilder.WithImage
            _builder.WithImage(image, callback)
            _builder.ShowImage()
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IButtonBuilder Implements IButtonBuilder.WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As IButtonBuilder Implements IButtonBuilder.WithImage
            _builder.WithImage(imagePath)
            _builder.ShowImage()
            Return Me
        End Function

        Public Function WithDescription(description As String) As IButtonBuilder Implements IButtonBuilder.WithDescription
            _builder.WithDescription(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithDescription
            _builder.WithDescription(description, callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As IButtonBuilder Implements IButtonBuilder.WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IButtonBuilder Implements IButtonBuilder.WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As IButtonBuilder Implements IButtonBuilder.InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As IButtonBuilder Implements IButtonBuilder.InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As IButtonBuilder Implements IButtonBuilder.InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As IButtonBuilder Implements IButtonBuilder.InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function ShowImage() As IButtonBuilder Implements IButtonBuilder.ShowImage
            _builder.ShowImage()
            Return Me
        End Function

        Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.ShowImage
            _builder.ShowImage(getShowImage)
            Return Me
        End Function

        Public Function HideImage() As IButtonBuilder Implements IButtonBuilder.HideImage
            _builder.HideImage()
            Return Me
        End Function

        Public Function HideImage(getShowImage As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.HideImage
            _builder.HideImage(getShowImage)
            Return Me
        End Function

		Public Function BeforeClick(callBack As OnAction, ParamArray actions() As Action(Of Button, Button.BeforeClickEventArgs)) As IButtonBuilder Implements IButtonBuilder.BeforeClick
			_builder.ThatDoes(callBack, actions.Select(Function(a) _builder.AsEventHandler(a)), beforeEvent:= True)
            Return Me
		End Function

        Public Function OnClick(callBack As OnAction, ParamArray actions() As Action(Of Button)) As IButtonBuilder Implements IButtonBuilder.OnClick
			_builder.ThatDoes(callBack, actions.Select(Function(a) _builder.AsEventHandler(a)), beforeEvent:= True)
            Return Me
		End Function

    End Class

End NameSpace