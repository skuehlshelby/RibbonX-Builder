Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders
    Public Class MenuBuilder
        Implements IInsert(Of MenuBuilder)
        Implements IEnabled(Of MenuBuilder)
        Implements IVisible(Of MenuBuilder)
        Implements ILabelScreenTipSuperTip(Of MenuBuilder)
        Implements IShowLabel(Of MenuBuilder)
        Implements IImage(Of MenuBuilder)
        Implements IDescription(Of MenuBuilder)
        Implements ISize(Of MenuBuilder)
        Implements IKeyTip(Of MenuBuilder)

        Private ReadOnly _builder As ControlBuilder

        'TODO Create Interface to set size of child menu items

        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Menu)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function Build(tag As Object) As Menu
            Return New Menu(_builder.Build(), tag:=tag)
        End Function

        Public Function Enabled() As MenuBuilder Implements IEnabled(Of MenuBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As MenuBuilder Implements IEnabled(Of MenuBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As MenuBuilder Implements IEnabled(Of MenuBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As MenuBuilder Implements IEnabled(Of MenuBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As MenuBuilder Implements IVisible(Of MenuBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As MenuBuilder Implements IVisible(Of MenuBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As MenuBuilder Implements IVisible(Of MenuBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As MenuBuilder Implements IVisible(Of MenuBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As MenuBuilder Implements ILabelScreenTipSuperTip(Of MenuBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As MenuBuilder Implements IShowLabel(Of MenuBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As MenuBuilder Implements IShowLabel(Of MenuBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As MenuBuilder Implements IShowLabel(Of MenuBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As MenuBuilder Implements IShowLabel(Of MenuBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(image)
            Return Me
        End Function

        Public Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As MenuBuilder Implements IImage(Of MenuBuilder).WithImage
            _builder.WithImage(imagePath)
            Return Me
        End Function

        Public Function WithDescription(description As String) As MenuBuilder Implements IDescription(Of MenuBuilder).WithDescription
            _builder.WithDescription(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As MenuBuilder Implements IDescription(Of MenuBuilder).WithDescription
            _builder.WithDescription(description, callback)
            Return Me
        End Function

        Public Function Large() As MenuBuilder Implements ISize(Of MenuBuilder).Large
            _builder.Large()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As MenuBuilder Implements ISize(Of MenuBuilder).Large
            _builder.Large(callback)
            Return Me
        End Function

        Public Function Normal() As MenuBuilder Implements ISize(Of MenuBuilder).Normal
            _builder.Normal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As MenuBuilder Implements ISize(Of MenuBuilder).Normal
            _builder.Normal(callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As MenuBuilder Implements IKeyTip(Of MenuBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As MenuBuilder Implements IKeyTip(Of MenuBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As MenuBuilder Implements IInsert(Of MenuBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace