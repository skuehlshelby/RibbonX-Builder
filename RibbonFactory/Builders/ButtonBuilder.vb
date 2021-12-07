Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders
    Public NotInheritable Class ButtonBuilder
        Implements IEnabled(Of ButtonBuilder)
        Implements IVisible(Of ButtonBuilder)
        Implements ILabelScreenTipSuperTip(Of ButtonBuilder)
        Implements IShowLabel(Of ButtonBuilder)
        Implements IShowImage(Of ButtonBuilder)
        Implements IOnActionClick(Of ButtonBuilder)
        Implements ISize(Of ButtonBuilder)
        Implements IImage(Of ButtonBuilder)
        Implements IDescription(Of ButtonBuilder)
        Implements IKeyTip(Of ButtonBuilder)
        Implements IInsert(Of ButtonBuilder)

        Private ReadOnly _builder As ControlBuilder

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Button)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As Button
            Return New Button(_builder.Build(), tag)
        End Function

        Public Function Enabled() As ButtonBuilder Implements IEnabled(Of ButtonBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As ButtonBuilder Implements IEnabled(Of ButtonBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As ButtonBuilder Implements IEnabled(Of ButtonBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As ButtonBuilder Implements IEnabled(Of ButtonBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As ButtonBuilder Implements IVisible(Of ButtonBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ButtonBuilder Implements IVisible(Of ButtonBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As ButtonBuilder Implements IVisible(Of ButtonBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ButtonBuilder Implements IVisible(Of ButtonBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As ButtonBuilder Implements IShowLabel(Of ButtonBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As ButtonBuilder Implements IShowLabel(Of ButtonBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As ButtonBuilder Implements IShowLabel(Of ButtonBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As ButtonBuilder Implements IShowLabel(Of ButtonBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function ThatDoes(callback As OnAction, action As Action) As ButtonBuilder Implements IOnActionClick(Of ButtonBuilder).ThatDoes
            _builder.ThatDoes(action, callback)
            Return Me
        End Function

        Public Function Large() As ButtonBuilder Implements ISize(Of ButtonBuilder).Large
            _builder.Large()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As ButtonBuilder Implements ISize(Of ButtonBuilder).Large
            _builder.Large(callback)
            Return Me
        End Function

        Public Function Normal() As ButtonBuilder Implements ISize(Of ButtonBuilder).Normal
            _builder.Normal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As ButtonBuilder Implements ISize(Of ButtonBuilder).Normal
            _builder.Normal(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As ButtonBuilder Implements IImage(Of ButtonBuilder).WithImage
            _builder.WithImage(image)
            _builder.ShowImage()
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As ButtonBuilder Implements IImage(Of ButtonBuilder).WithImage
            _builder.WithImage(image, callback)
            _builder.ShowImage()
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As ButtonBuilder Implements IImage(Of ButtonBuilder).WithImage
            _builder.WithImage(imagePath)
            _builder.ShowImage()
            Return Me
        End Function

        Public Function WithDescription(description As String) As ButtonBuilder Implements IDescription(Of ButtonBuilder).WithDescription
            _builder.WithDescription(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As ButtonBuilder Implements IDescription(Of ButtonBuilder).WithDescription
            _builder.WithDescription(description, callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As ButtonBuilder Implements IKeyTip(Of ButtonBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ButtonBuilder Implements IKeyTip(Of ButtonBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ButtonBuilder Implements IInsert(Of ButtonBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ButtonBuilder Implements IInsert(Of ButtonBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ButtonBuilder Implements IInsert(Of ButtonBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ButtonBuilder Implements IInsert(Of ButtonBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function ShowImage() As ButtonBuilder Implements IShowImage(Of ButtonBuilder).ShowImage
            _builder.ShowImage()
            Return Me
        End Function

        Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As ButtonBuilder Implements IShowImage(Of ButtonBuilder).ShowImage
            _builder.ShowImage(getShowImage)
            Return Me
        End Function

        Public Function HideImage() As ButtonBuilder Implements IShowImage(Of ButtonBuilder).HideImage
            _builder.HideImage()
            Return Me
        End Function

        Public Function HideImage(getShowImage As FromControl(Of Boolean)) As ButtonBuilder Implements IShowImage(Of ButtonBuilder).HideImage
            _builder.HideImage(getShowImage)
            Return Me
        End Function

    End Class

End NameSpace