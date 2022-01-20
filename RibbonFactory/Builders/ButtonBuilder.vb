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
        Implements IBuilder(Of Button)
        Implements IID(Of ButtonBuilder)
        Implements IEnabled(Of ButtonBuilder)
        Implements IVisible(Of ButtonBuilder)
        Implements ILabelScreenTipSuperTip(Of ButtonBuilder)
        Implements IShowLabel(Of ButtonBuilder)
        Implements IShowImage(Of ButtonBuilder)
        Implements IOnActionClick(Of Button, ButtonBuilder)
        Implements ISize(Of ButtonBuilder)
        Implements IImage(Of ButtonBuilder)
        Implements IDescription(Of ButtonBuilder)
        Implements IKeyTip(Of ButtonBuilder)
        Implements IInsert(Of ButtonBuilder)

        Private ReadOnly _builder As ControlBuilder
        Private ReadOnly _beforeClickHandlers As ICollection(Of EventHandler(Of Button.BeforeClickEventArgs)) = New List(Of EventHandler(Of Button.BeforeClickEventArgs))
		Private ReadOnly _clickHandlers As ICollection(Of EventHandler) = New List(Of EventHandler)

        Public Sub New()
            Dim defaultProvider As IDefaultProvider = New DefaultProvider(Of Button)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(defaultProvider)
            _builder = New ControlBuilder(attributeGroupBuilder)
        End Sub

        Public Sub New(template As Button)
            Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder()
            attributeGroupBuilder.SetDefaults(template)
            attributeGroupBuilder.AddID(IdManager.GetID(Of Button)())
            _builder = New ControlBuilder(attributeGroupBuilder)

            Array.ForEach(template.GetBeforeClickInvocationList(), Sub(handler) _beforeClickHandlers.Add(handler))
			Array.ForEach(template.GetClickInvocationList(), Sub(handler) _clickHandlers.Add(handler))
        End Sub

        Public Sub New(template As RibbonElement)
            Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

            If defaultProvider IsNot Nothing Then
                Dim templateAttributes As AttributeSet = defaultProvider.GetDefaults()
                Dim buttonAttributes As AttributeSet = New DefaultProvider(Of Button)().GetDefaults()
                Dim intersection As AttributeSet = New AttributeSet(templateAttributes.Where(Function(a) buttonAttributes.Contains(a)))
                Dim attributeGroupBuilder As AttributeGroupBuilder = New AttributeGroupBuilder(intersection)
                attributeGroupBuilder.AddID(IdManager.GetID(Of Button)())
                _builder = New ControlBuilder(attributeGroupBuilder)
            Else
                Throw New ArgumentException($"Could not copy applicable properties of type '{template.GetType().Name}' to type '{GetType(Button)}'")
            End If
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As Button Implements IBuilder(Of Button).Build

            Dim button As Button = New Button(_builder.Build(), tag)

            For Each handler As EventHandler(Of Button.BeforeClickEventArgs) In _beforeClickHandlers
				AddHandler button.BeforeClick, handler
			Next

			For Each handler As EventHandler In _clickHandlers
				AddHandler button.OnClick, handler
			Next

            Return button
        End Function

        Public Function WithId(id As String) As ButtonBuilder Implements IID(Of ButtonBuilder).WithId
            _builder.WithId(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As ButtonBuilder Implements IID(Of ButtonBuilder).WithIdQ
            _builder.WithId([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As ButtonBuilder Implements IID(Of ButtonBuilder).WithIdMso
            _builder.WithId(mso)
            Return Me
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
            _builder.ShowLabel()
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ILabelScreenTipSuperTip(Of ButtonBuilder).WithLabel
            _builder.WithLabel(label, callback, copyToScreenTip)
            _builder.ShowLabel()
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

        Public Function ThatDoes(action As Action(Of Button), callback As OnAction) As ButtonBuilder Implements IOnActionClick(Of Button, ButtonBuilder).ThatDoes
            _builder.ThatDoes(callback)
            
            With New OnClickHelper(action)
                _clickHandlers.Add(AddressOf .Handle)
            End With

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

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As ButtonBuilder Implements IImage(Of ButtonBuilder).WithImage
            _builder.WithImage(image, callback)
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

        Public Function BeforeClick(callback As OnAction, Paramarray actions() As Action(Of Button, Button.BeforeClickEventArgs)) As ButtonBuilder
            For Each action As Action(Of Button, Button.BeforeClickEventArgs) In actions
                With New BeforeClickHelper(action)
                    _beforeClickHandlers.Add(AddressOf .Handle)
                End With
            Next

            _builder.ThatDoes(callback)

            Return Me
        End Function

        Public Function OnClick(callback As OnAction, Paramarray actions() As Action(Of Button)) As ButtonBuilder
            For Each action As Action(Of Button) In actions
                With New OnClickHelper(action)
                    _clickHandlers.Add(AddressOf .Handle)
                End With
            Next

            _builder.ThatDoes(callback)

            Return Me
        End Function

        Private NotInheritable Class BeforeClickHelper
            Private Readonly action As Action(Of Button, Button.BeforeClickEventArgs)

			Public Sub New(action As Action(Of Button, Button.BeforeClickEventArgs))
				Me.action = action
			End Sub

            Public Sub Handle(sender As Object, e As Button.BeforeClickEventArgs)
                action.Invoke(DirectCast(sender, Button), e)
            End Sub

		End Class

        Private NotInheritable Class OnClickHelper
            Private Readonly action As Action(Of Button)

			Public Sub New(action As Action(Of Button))
				Me.action = action
			End Sub

            Public Sub Handle(sender As Object, e As EventArgs)
                action.Invoke(DirectCast(sender, Button))
            End Sub

		End Class

    End Class

End NameSpace