Imports System.Drawing
Imports RibbonFactory.Builders
Imports RibbonFactory.Controls
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities
Imports stdole

Public NotInheritable Class RandomControlGenerator

    Private ReadOnly logs As Dictionary(Of RibbonElement, Dictionary(Of String, Object)) = New Dictionary(Of RibbonElement, Dictionary(Of String, Object))()
    Private ReadOnly random As Random = New Random()
    Private ReadOnly callbacks As ICreateCallbacks

    Private Class AttributeSetInjector
        Inherits RibbonElement
        Implements IDefaultProvider

        Private ReadOnly attributes As AttributeSet

        Public Sub New(attributes As AttributeSet)
            Me.attributes = attributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Throw New NotSupportedException()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Throw New NotSupportedException()
            End Get
        End Property

        Public Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return attributes
        End Function

    End Class

    Private Class FakeBuilder
        Inherits BuilderBase(Of RibbonElement)

        Private ReadOnly parent As RandomControlGenerator
        Public ReadOnly Property CreationReport As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()

        Public Sub New(parent As RandomControlGenerator)
            Me.parent = parent
        End Sub

        Public Function GetRandomControl() As AttributeSetInjector
            MaybeDo(Sub() ColumnCountBase(parent.random.Next(0, 20)))
            MaybeDo(Sub() RowCountBase(parent.random.Next(0, 20)))
            MaybeDo(Sub() DisabledBase(), Sub() EnabledBase(), Sub() DisabledBase(AddressOf parent.callbacks.GetEnabled), Sub() EnabledBase(AddressOf parent.callbacks.GetEnabled))
            MaybeDo(Sub() InvisibleBase(), Sub() VisibleBase(), Sub() InvisibleBase(AddressOf parent.callbacks.GetVisible), Sub() VisibleBase(AddressOf parent.callbacks.GetVisible))
            MaybeDo(Sub() DescriptionBase("The Description"), Sub() DescriptionBase("The Description", AddressOf parent.callbacks.GetDescription))
            MaybeDo(Sub() LabelBase("The Label"), Sub() LabelBase("The Label", AddressOf parent.callbacks.GetLabel))
            MaybeDo(Sub() ScreentipBase("The Screentip"), Sub() ScreentipBase("The Screentip", AddressOf parent.callbacks.GetScreenTip))
            MaybeDo(Sub() SupertipBase("The Supertip"), Sub() SupertipBase("The Supertip", AddressOf parent.callbacks.GetSuperTip))
            MaybeDo(Sub() KeyTipBase("A5"), Sub() KeyTipBase("A5", AddressOf parent.callbacks.GetKeyTip))
            MaybeDo(Sub() TitleBase("The Title")) 'TODO Add GetTitle to ICreateCallbacks
            MaybeDo(Sub() DoNotShowInRibbonBase())
            MaybeDo(Sub() HideImageBase(), Sub() ShowImageBase(), Sub() HideImageBase(AddressOf parent.callbacks.GetShowImage), Sub() ShowImageBase(AddressOf parent.callbacks.GetShowImage))
            MaybeDo(Sub() HideItemImageBase(), Sub() ShowItemImageBase())
            MaybeDo(Sub() HideItemLabelBase(), Sub() ShowItemLabelBase())
            MaybeDo(Sub() HorizontalBase(), Sub() VerticalBase())
            MaybeDo(Sub() NormalBase(), Sub() LargeBase(), Sub() NormalBase(AddressOf parent.callbacks.GetSize), Sub() LargeBase(AddressOf parent.callbacks.GetSize))
            MaybeDo(Sub() ItemHeightBase(parent.random.Next(1, 40)), Sub() ItemHeightBase(parent.random.Next(1, 40), AddressOf parent.callbacks.GetItemHeight))
            MaybeDo(Sub() ItemWidthBase(parent.random.Next(1, 40)), Sub() ItemWidthBase(parent.random.Next(1, 40), AddressOf parent.callbacks.GetItemWidth))
            MaybeDo(Sub() InvalidateContentOnDropBase())
            MaybeDo(Sub() MaxLengthBase(CByte(parent.random.Next(0, 255))))
            MaybeDo(Sub() TextBase("Some Text", AddressOf parent.callbacks.GetText, AddressOf parent.callbacks.OnChange))
            MaybeDo(Sub() PressedBase(True, AddressOf parent.callbacks.GetPressed, AddressOf parent.callbacks.OnButtonToggle))

            Return New AttributeSetInjector(Build())
        End Function

#Region "Enabled"

        Protected Shadows Sub EnabledBase()
            creationReport.Add(NameOf(IEnabled.Enabled), True)
            MyBase.EnabledBase()
        End Sub

        Protected Shadows Sub EnabledBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IEnabled.Enabled), True)
            MyBase.EnabledBase(callback)
        End Sub

#End Region

#Region "Disabled"

        Protected Shadows Sub DisabledBase()
            creationReport.Add(NameOf(IEnabled.Enabled), False)
            MyBase.DisabledBase()
        End Sub

        Protected Shadows Sub DisabledBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IEnabled.Enabled), False)
            MyBase.DisabledBase(callback)
        End Sub

#End Region

#Region "Visible"
        Protected Shadows Sub VisibleBase()
            creationReport.Add(NameOf(IVisible.Visible), True)
            MyBase.VisibleBase()
        End Sub

        Protected Shadows Sub VisibleBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IVisible.Visible), True)
            MyBase.VisibleBase(callback)
        End Sub
#End Region

#Region "Invisible"
        Protected Shadows Sub InvisibleBase()
            creationReport.Add(NameOf(IVisible.Visible), False)
            MyBase.InvisibleBase()
        End Sub

        Protected Shadows Sub InvisibleBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IVisible.Visible), False)
            MyBase.InvisibleBase(callback)
        End Sub
#End Region

#Region "Id"

        Protected Shadows Sub WithIdBase(id As String)
            creationReport.Add(NameOf(RibbonElement.ID), id)
            MyBase.WithIdBase(id)
        End Sub

        Protected Shadows Sub WithIdBase([namespace] As String, id As String)
            creationReport.Add(NameOf(RibbonElement.ID), $"{[namespace]}:{id}")
            MyBase.WithIdBase([namespace], id)
        End Sub

        Protected Shadows Sub WithIdBase(mso As MSO)
            creationReport.Add(NameOf(RibbonElement.ID), mso)
            MyBase.WithIdBase(mso)
        End Sub

#End Region

#Region "Label"

        Protected Shadows Sub LabelBase(label As String)
            creationReport.Add(NameOf(ILabel.Label), label)
            MyBase.LabelBase(label)
        End Sub

        Protected Shadows Sub LabelBase(label As String, callback As FromControl(Of String))
            creationReport.Add(NameOf(ILabel.Label), label)
            MyBase.LabelBase(label, callback)
        End Sub

#End Region

#Region "Show Label"

        Protected Shadows Sub HideLabelBase()
            creationReport.Add(NameOf(IShowLabel.ShowLabel), False)
            MyBase.HideLabelBase()
        End Sub

        Protected Shadows Sub HideLabelBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IShowLabel.ShowLabel), False)
            MyBase.HideLabelBase(callback)
        End Sub

        Protected Shadows Sub ShowLabelBase()
            creationReport.Add(NameOf(IShowLabel.ShowLabel), True)
            MyBase.ShowLabelBase()
        End Sub

        Protected Shadows Sub ShowLabelBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IShowLabel.ShowLabel), True)
            MyBase.ShowLabelBase(callback)
        End Sub

#End Region

#Region "Description"

        Protected Shadows Sub DescriptionBase(description As String)
            creationReport.Add(NameOf(IDescription.Description), description)
            MyBase.DescriptionBase(description)
        End Sub

        Protected Shadows Sub DescriptionBase(description As String, callback As FromControl(Of String))
            creationReport.Add(NameOf(IDescription.Description), description)
            MyBase.DescriptionBase(description, callback)
        End Sub

#End Region

#Region "KeyTip"

        Protected Shadows Sub KeyTipBase(keyTip As KeyTip)
            creationReport.Add(NameOf(IKeyTip.KeyTip), keyTip)
            MyBase.KeyTipBase(keyTip)
        End Sub

        Protected Shadows Sub KeyTipBase(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            creationReport.Add(NameOf(IKeyTip.KeyTip), keyTip)
            MyBase.KeyTipBase(keyTip, callback)
        End Sub

#End Region

#Region "Screentip"

        Protected Shadows Sub ScreentipBase(screentip As String)
            creationReport.Add(NameOf(IScreenTip.ScreenTip), screentip)
            MyBase.ScreentipBase(screentip)
        End Sub

        Protected Shadows Sub ScreentipBase(screentip As String, callback As FromControl(Of String))
            creationReport.Add(NameOf(IScreenTip.ScreenTip), screentip)
            MyBase.ScreentipBase(screentip, callback)
        End Sub

#End Region

#Region "Supertip"

        Protected Shadows Sub SupertipBase(supertip As String)
            creationReport.Add(NameOf(ISuperTip.SuperTip), supertip)
            MyBase.SupertipBase(supertip)
        End Sub

        Protected Shadows Sub SupertipBase(supertip As String, callback As FromControl(Of String))
            creationReport.Add(NameOf(ISuperTip.SuperTip), supertip)
            MyBase.SupertipBase(supertip, callback)
        End Sub

#End Region

#Region "Title"

        Protected Shadows Sub TitleBase(title As String)
            creationReport.Add(NameOf(ITitle.Title), title)
            MyBase.TitleBase(title)
        End Sub

        Protected Shadows Sub TitleBase(title As String, callback As FromControl(Of String))
            creationReport.Add(NameOf(ITitle.Title), title)
            MyBase.TitleBase(title, callback)
        End Sub

#End Region

#Region "Image"

        Protected Shadows Sub ImageBase(path As String)
            creationReport.Add(NameOf(IImage.Image), path)
            MyBase.ImageBase(path)
        End Sub

        Protected Shadows Sub ImageBase(image As ImageMSO)
            creationReport.Add(NameOf(IImage.Image), image)
            MyBase.ImageBase(image)
        End Sub

        Protected Shadows Sub ImageBase(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
            creationReport.Add(NameOf(IImage.Image), image)
            MyBase.ImageBase(image, callback)
        End Sub

        Protected Shadows Sub ImageBase(image As Bitmap, callback As FromControl(Of IPictureDisp))
            ImageBase(New AdapterForIPicture(image), callback)
        End Sub

        Protected Shadows Sub ImageBase(image As Icon, callback As FromControl(Of IPictureDisp))
            ImageBase(New AdapterForIPicture(image), callback)
        End Sub

#End Region

#Region "Show Image"

        Protected Shadows Sub ShowImageBase()
            creationReport.Add(NameOf(IShowImage.ShowImage), True)
            MyBase.ShowImageBase()
        End Sub

        Protected Shadows Sub ShowImageBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IShowImage.ShowImage), True)
            MyBase.ShowImageBase(callback)
        End Sub

#End Region

#Region "Hide Image"

        Protected Shadows Sub HideImageBase()
            creationReport.Add(NameOf(IShowImage.ShowImage), False)
            MyBase.HideImageBase()
        End Sub

        Protected Shadows Sub HideImageBase(callback As FromControl(Of Boolean))
            creationReport.Add(NameOf(IShowImage.ShowImage), False)
            MyBase.HideImageBase(callback)
        End Sub

#End Region

#Region "Normal"

        Protected Shadows Sub NormalBase()
            creationReport.Add(NameOf(ISize.Size), ControlSize.normal)
            MyBase.NormalBase()
        End Sub

        Protected Shadows Sub NormalBase(callback As FromControl(Of ControlSize))
            creationReport.Add(NameOf(ISize.Size), ControlSize.normal)
            MyBase.NormalBase(callback)
        End Sub

#End Region

#Region "Large"

        Protected Shadows Sub LargeBase()
            creationReport.Add(NameOf(ISize.Size), ControlSize.large)
            MyBase.LargeBase()
        End Sub

        Protected Shadows Sub LargeBase(callback As FromControl(Of ControlSize))
            creationReport.Add(NameOf(ISize.Size), ControlSize.large)
            MyBase.LargeBase(callback)
        End Sub

#End Region

#Region "Max Length"

        Protected Shadows Sub MaxLengthBase(maxLength As Byte)
            creationReport.Add(NameOf(IMaxLength.MaxLength), maxLength)
            MyBase.MaxLengthBase(maxLength)
        End Sub

#End Region

#Region "Horizontal/Vertical"

        Protected Shadows Sub HorizontalBase()
            creationReport.Add(NameOf(IBoxStyle.BoxStyle), BoxStyle.horizontal)
            MyBase.HorizontalBase()
        End Sub

        Protected Shadows Sub VerticalBase()
            creationReport.Add(NameOf(IBoxStyle.BoxStyle), BoxStyle.vertical)
            MyBase.VerticalBase()
        End Sub

#End Region

#Region "Get/Set Text"

        Protected Shadows Sub TextBase(text As String, getText As FromControl(Of String), setText As TextChanged)
            creationReport.Add(NameOf(IText.Text), text)
            MyBase.TextBase(text, getText, setText)
        End Sub

#End Region

#Region "Get/Set Pressed"

        Protected Shadows Sub PressedBase(pressed As Boolean, getPressed As FromControl(Of Boolean), setPressed As ButtonPressed)
            creationReport.Add(NameOf(IPressed.Checked), pressed)
            MyBase.PressedBase(pressed, getPressed, setPressed)
        End Sub

#End Region

        '' Child Item Methods ''

#Region "Item Height/Width"

        Protected Shadows Sub ItemHeightBase(height As Integer)
            creationReport.Add(NameOf(IItemDimensions.ItemHeight), height)
            MyBase.ItemHeightBase(height)
        End Sub

        Protected Shadows Sub ItemHeightBase(height As Integer, callback As FromControl(Of Integer))
            creationReport.Add(NameOf(IItemDimensions.ItemHeight), height)
            MyBase.ItemHeightBase(height, callback)
        End Sub

        Protected Shadows Sub ItemWidthBase(width As Integer)
            creationReport.Add(NameOf(IItemDimensions.ItemWidth), width)
            MyBase.ItemWidthBase(width)
        End Sub

        Protected Shadows Sub ItemWidthBase(width As Integer, callback As FromControl(Of Integer))
            creationReport.Add(NameOf(IItemDimensions.ItemWidth), width)
            MyBase.ItemWidthBase(width, callback)
        End Sub

#End Region

        Private Sub MaybeDo(action As Action)
            If parent.random.NextBoolean() Then
                action.Invoke()
            End If
        End Sub

        Private Sub MaybeDo(ParamArray actions() As Action)
            If parent.random.NextBoolean() Then
                actions(parent.random.Next(actions.Length)).Invoke()
            End If
        End Sub
    End Class

    Public Sub New(callbacks As ICreateCallbacks)
        Me.callbacks = callbacks
    End Sub

    Private Function MakeControl(Of T As RibbonElement)(creationMethod As Func(Of RibbonElement, T)) As T
        Dim builder As FakeBuilder = New FakeBuilder(Me)
        Dim control As T = creationMethod.Invoke(builder.GetRandomControl())
        logs.Add(control, builder.CreationReport)

        Return control
    End Function

    Public Function MakeButton() As Button
        Return MakeControl(Function(e) New Button(Sub() Return, e))
    End Function

    Public Function MakeCheckbox() As CheckBox
        Return MakeControl(Function(e) New CheckBox(Sub() Return, e))
    End Function

    Public Function MakeComboBox() As ComboBox
        Return MakeControl(Function(e) New ComboBox(Sub() Return, e))
    End Function

    Public Function MakeEditBox() As EditBox
        Return MakeControl(Function(e) New EditBox(e, Sub() Return))
    End Function

    Public Function MakeLabelControl() As LabelControl
        Return MakeControl(Function(e) New LabelControl(e, Sub() Return))
    End Function

    Public Function MakeToggleButton() As ToggleButton
        Return MakeControl(Function(e) New ToggleButton(e, Sub() Return))
    End Function

    Public Function MakeSplitButtonWithRegularButtonAndMenu() As SplitButton
        Return New SplitButton(New FakeBuilder(Me).GetRandomControl(), Sub() Return, MakeButton(), MakeMenu())
    End Function

    Public Function MakeDropDown() As DropDown
        Return MakeControl(Function(e) New DropDown(Sub() Return, e))
    End Function

    Public Function MakeGallery() As Gallery
        Return MakeControl(Function(e) New Gallery(Sub() Return, Nothing, e))
    End Function

    Public Function MakeMenu() As Menu
        Return MakeControl(Function(e) New Menu(e, Sub() Return, Nothing))
    End Function

    Public Function MakeBox(ParamArray controls() As RibbonElement) As Box
        Return MakeControl(Function(e) New Box(e, Sub() Return, controls))
    End Function

    Public Function MakeGroup(ParamArray controls() As RibbonElement) As Group
        Return MakeControl(Function(e) New Group(Sub() Return, controls, e))
    End Function

    Public Function MakeTab(ParamArray groups() As Group) As Tab
        Return MakeControl(Function(e) New Tab(Sub() Return, groups, e))
    End Function

    Public Function MakeRibbon(ParamArray tabs() As Tab) As Ribbon
        Return New Ribbon(tabs)
    End Function

    Public Function GetCreationLogs() As Dictionary(Of RibbonElement, Dictionary(Of String, Object))
        Return logs
    End Function

    Public Sub ClearLogs()
        logs.Clear()
    End Sub

End Class