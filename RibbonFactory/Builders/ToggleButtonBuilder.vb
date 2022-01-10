Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Builders

    Public Class ToggleButtonBuilder
        Implements IBuilder(Of ToggleButton)
        Implements IID(Of ToggleButtonBuilder)
        Implements IVisible(Of ToggleButtonBuilder)
        Implements IEnabled(Of ToggleButtonBuilder)
        Implements IInsert(Of ToggleButtonBuilder)
        Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder)
        Implements IShowLabel(Of ToggleButtonBuilder)
        Implements IKeyTip(Of ToggleButtonBuilder)
        Implements IImage(Of ToggleButtonBuilder)
        Implements IShowImage(Of ToggleButtonBuilder)
        Implements IDescription(Of ToggleButtonBuilder)
        Implements ISize(Of ToggleButtonBuilder)
        Implements IOnActionToggle(Of ToggleButtonBuilder)

        Private ReadOnly _builder As ControlBuilder

        Friend Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of SplitButton)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As ToggleButton)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of ToggleButton)())
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As RibbonElement)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim toggleButtonAttributes As AttributeSet = New DefaultProvider(Of ToggleButton)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) toggleButtonAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of ToggleButton)())
                _builder = New ControlBuilder(attributeGroupBuilder)
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(ToggleButton)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As ToggleButton Implements IBuilder(Of ToggleButton).Build
            Return New ToggleButton(_builder.Build(), tag:=tag)
        End Function

        Public Function WithId(id As String) As ToggleButtonBuilder Implements IID(Of ToggleButtonBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As ToggleButtonBuilder Implements IID(Of ToggleButtonBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As ToggleButtonBuilder Implements IID(Of ToggleButtonBuilder).WithIdMso
            _builder.WithId(mso)
            Return Me
        End Function

        Public Function Enabled() As ToggleButtonBuilder Implements IEnabled(Of ToggleButtonBuilder).Enabled
            _builder.Enabled()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IEnabled(Of ToggleButtonBuilder).Enabled
            _builder.Enabled(callback)
            Return Me
        End Function

        Public Function Disabled() As ToggleButtonBuilder Implements IEnabled(Of ToggleButtonBuilder).Disabled
            _builder.Disabled()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IEnabled(Of ToggleButtonBuilder).Disabled
            _builder.Disabled(callback)
            Return Me
        End Function

        Public Function Visible() As ToggleButtonBuilder Implements IVisible(Of ToggleButtonBuilder).Visible
            _builder.Visible()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IVisible(Of ToggleButtonBuilder).Visible
            _builder.Visible(callback)
            Return Me
        End Function

        Public Function Invisible() As ToggleButtonBuilder Implements IVisible(Of ToggleButtonBuilder).Invisible
            _builder.Invisible()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IVisible(Of ToggleButtonBuilder).Invisible
            _builder.Invisible(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ToggleButtonBuilder Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder).WithLabel
            _builder.WithLabel(label, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ToggleButtonBuilder Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel() As ToggleButtonBuilder Implements IShowLabel(Of ToggleButtonBuilder).ShowLabel
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IShowLabel(Of ToggleButtonBuilder).ShowLabel
            _builder.ShowLabel(callback)
            Return Me
        End Function

        Public Function HideLabel() As ToggleButtonBuilder Implements IShowLabel(Of ToggleButtonBuilder).HideLabel
            _builder.HideLabel()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IShowLabel(Of ToggleButtonBuilder).HideLabel
            _builder.HideLabel(callback)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ToggleButtonBuilder Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder).WithScreenTip
            _builder.WithScreentip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ToggleButtonBuilder Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder).WithScreenTip
            _builder.WithScreentip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ToggleButtonBuilder Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder).WithSuperTip
            _builder.WithSupertip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ToggleButtonBuilder Implements ILabelScreenTipSuperTip(Of ToggleButtonBuilder).WithSuperTip
            _builder.WithSupertip(superTip, callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As ToggleButtonBuilder Implements IKeyTip(Of ToggleButtonBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ToggleButtonBuilder Implements IKeyTip(Of ToggleButtonBuilder).WithKeyTip
            _builder.WithKeyTip(keyTip, callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As ToggleButtonBuilder Implements IImage(Of ToggleButtonBuilder).WithImage
            _builder.WithImage(image)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As ToggleButtonBuilder Implements IImage(Of ToggleButtonBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As ToggleButtonBuilder Implements IImage(Of ToggleButtonBuilder).WithImage
            _builder.WithImage(image, callback)
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As ToggleButtonBuilder Implements IImage(Of ToggleButtonBuilder).WithImage
            _builder.WithImage(imagePath)
            Return Me
        End Function

        Public Function WithDescription(description As String) As ToggleButtonBuilder Implements IDescription(Of ToggleButtonBuilder).WithDescription
            _builder.WithDescription(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As ToggleButtonBuilder Implements IDescription(Of ToggleButtonBuilder).WithDescription
            _builder.WithDescription(description, callback)
            Return Me
        End Function

        Public Function Large() As ToggleButtonBuilder Implements ISize(Of ToggleButtonBuilder).Large
            _builder.Large()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As ToggleButtonBuilder Implements ISize(Of ToggleButtonBuilder).Large
            _builder.Large(callback)
            Return Me
        End Function

        Public Function Normal() As ToggleButtonBuilder Implements ISize(Of ToggleButtonBuilder).Normal
            _builder.Normal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As ToggleButtonBuilder Implements ISize(Of ToggleButtonBuilder).Normal
            _builder.Normal(callback)
            Return Me
        End Function

        Public Function ThatDoes(action As Action, callback As ButtonPressed) As ToggleButtonBuilder Implements IOnActionToggle(Of ToggleButtonBuilder).ThatDoes
            _builder.ThatDoes(action, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ToggleButtonBuilder Implements IInsert(Of ToggleButtonBuilder).InsertBefore
            _builder.InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ToggleButtonBuilder Implements IInsert(Of ToggleButtonBuilder).InsertBefore
            _builder.InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ToggleButtonBuilder Implements IInsert(Of ToggleButtonBuilder).InsertAfter
            _builder.InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ToggleButtonBuilder Implements IInsert(Of ToggleButtonBuilder).InsertAfter
            _builder.InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function ShowImage() As ToggleButtonBuilder Implements IShowImage(Of ToggleButtonBuilder).ShowImage
            _builder.ShowImage()
            Return Me
        End Function

        Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IShowImage(Of ToggleButtonBuilder).ShowImage
            _builder.ShowImage(getShowImage)
            Return Me
        End Function

        Public Function HideImage() As ToggleButtonBuilder Implements IShowImage(Of ToggleButtonBuilder).HideImage
            _builder.HideImage()
            Return Me
        End Function

        Public Function HideImage(getShowImage As FromControl(Of Boolean)) As ToggleButtonBuilder Implements IShowImage(Of ToggleButtonBuilder).HideImage
            _builder.HideImage(getShowImage)
            Return Me
        End Function

    End Class

End NameSpace