Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders
    Public NotInheritable Class GroupBuilder
        Implements IID(Of GroupBuilder)
        Implements IInsert(Of GroupBuilder)
        Implements IVisible(Of GroupBuilder)
        Implements ILabelScreenTipSuperTip(Of GroupBuilder)
        Implements IKeyTip(Of GroupBuilder)
        Implements IImage(Of GroupBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _controls As ICollection(Of RibbonElement)

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Group)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = new ControlBuilder(attributeGroupBuilder)
            _controls = New List(Of RibbonElement)
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As Group
            Return New Group(_controls, _builder.Build(), tag)
        End Function

        Public Function WithId(id As String) As GroupBuilder Implements IID(Of GroupBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As GroupBuilder Implements IID(Of GroupBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As GroupBuilder Implements IID(Of GroupBuilder).WithIdMso
            _builder.WithId(mso)
            Return Me
        End Function

        Public Function Visible() As GroupBuilder Implements IVisible(Of GroupBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As GroupBuilder Implements IVisible(Of GroupBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As GroupBuilder Implements IVisible(Of GroupBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As GroupBuilder Implements IVisible(Of GroupBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As GroupBuilder Implements IKeyTip(Of GroupBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As GroupBuilder Implements IKeyTip(Of GroupBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As GroupBuilder Implements IImage(Of GroupBuilder).WithImage
            _builder.WithImage(image)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As GroupBuilder Implements IImage(Of GroupBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As GroupBuilder Implements IImage(Of GroupBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As GroupBuilder Implements IImage(Of GroupBuilder).WithImage
            _builder.WithImage(imagePath)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As GroupBuilder Implements IInsert(Of GroupBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As GroupBuilder Implements IInsert(Of GroupBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As GroupBuilder Implements IInsert(Of GroupBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As GroupBuilder Implements IInsert(Of GroupBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As GroupBuilder Implements ILabelScreenTipSuperTip(Of GroupBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As GroupBuilder Implements ILabelScreenTipSuperTip(Of GroupBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As GroupBuilder Implements ILabelScreenTipSuperTip(Of GroupBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As GroupBuilder Implements ILabelScreenTipSuperTip(Of GroupBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As GroupBuilder Implements ILabelScreenTipSuperTip(Of GroupBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As GroupBuilder Implements ILabelScreenTipSuperTip(Of GroupBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function WithControl(control As RibbonElement) As GroupBuilder
            _controls.Add(control)
            Return Me
        End Function

        Public Function WithControls(ParamArray controls As RibbonElement()) As GroupBuilder
            For Each control As RibbonElement In controls
                _controls.Add(control)
            Next
            
            Return Me
        End Function

    End Class

End NameSpace