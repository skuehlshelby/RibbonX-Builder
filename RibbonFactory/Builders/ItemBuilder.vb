Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders
    Public NotInheritable Class ItemBuilder
        Implements ILabelScreenTipSuperTip(Of ItemBuilder)
        Implements IImage(Of ItemBuilder)

        Private ReadOnly _builder As ControlBuilder

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Item)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Function Build() As Item
            Return New Item(_builder.Build())
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ItemBuilder Implements ILabelScreenTipSuperTip(Of ItemBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ItemBuilder Implements ILabelScreenTipSuperTip(Of ItemBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ItemBuilder Implements ILabelScreenTipSuperTip(Of ItemBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ItemBuilder Implements ILabelScreenTipSuperTip(Of ItemBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ItemBuilder Implements ILabelScreenTipSuperTip(Of ItemBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ItemBuilder Implements ILabelScreenTipSuperTip(Of ItemBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As ItemBuilder Implements IImage(Of ItemBuilder).WithImage
            _builder.WithImage(imagePath)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As ItemBuilder Implements IImage(Of ItemBuilder).WithImage
            _builder.WithImage(image)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As ItemBuilder Implements IImage(Of ItemBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As ItemBuilder Implements IImage(Of ItemBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function
    End Class
End NameSpace