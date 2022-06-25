Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports RibbonX
Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.ComTypes.Extensibility
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.Images.BuiltIn
Imports Excel = Microsoft.Office.Interop.Excel


<ComVisible(True)>
<Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4")>
<ProgId("ExampleRibbon.Ribbon")>
Public Class Ribbon
    Inherits StockRibbonBase
    Implements IDTExtensibility2
    Implements IRibbonExtensibility

    Public Sub New()
        MyBase.New(New Troubleshooter())
    End Sub

    Private ReadOnly Property Excel As Excel.Application
        Get
            Return DirectCast(HostApp, Excel.Application)
        End Get
    End Property

    Protected Overrides Function BuildRibbon() As Controls.Ribbon
        Dim buttonWithStockIconOne As Button = New Button(
                Sub(bb) bb.
                Large().
                WithLabel("Happy Button").
                WithSuperTip("Oh, to be so happy again!").
                WithImage(Common.HappyFace).
                RouteClickTo(AddressOf OnAction).
                OnClick(Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")))

        Dim buttonWithStockIconTwo As Button = New Button(Sub(bb) bb.
                WithLabel("Sad Button").
                WithSuperTip("#Sad").
                WithImage(Common.SadFace),
                template:=buttonWithStockIconOne)

        Dim buttonWithStockIconThree As Button = New Button(Sub(bb) bb.
                WithLabel("Money Button").
                WithSuperTip("Make that money!").
                WithImage(Common.DollarSign),
                template:=buttonWithStockIconOne)


        Dim buttonWithStockIconOneSmall As Button = New Button(Sub(bb) bb.Normal(),
            template:=buttonWithStockIconOne)

        Dim buttonWithStockIconTwoSmall As Button = New Button(Sub(bb) bb.Normal(),
            template:=buttonWithStockIconTwo)

        Dim buttonWithStockIconThreeSmall As Button = New Button(Sub(bb) bb.Normal(),
            template:=buttonWithStockIconThree)

        Dim buttonsWithStockIcons As Group = New Group(Sub(gb) gb.
            WithLabel("Buttons With Stock Icons"),
            {buttonWithStockIconOne, buttonWithStockIconTwo, buttonWithStockIconThree, Separator.CreateVisible(), buttonWithStockIconOneSmall, buttonWithStockIconTwoSmall, buttonWithStockIconThreeSmall})

        Dim buttonWithCustomIconOne As Button = New Button(Sub(bb) bb.
            Large().
            WithLabel("GitHub").
            WithSuperTip("Open GitHub's website in the default browser.").
            WithImage(LoadBitmap("ExampleRibbon.github.png"), AddressOf GetImage).
            OnClick(Sub(b) OpenWebsiteInDefaultBrowser(b.Tag.ToString()), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")).
            RouteClickTo(AddressOf OnAction),
            tag:="https://github.com/")

        Dim buttonWithCustomIconTwo As Button = New Button(Sub(bb) bb.
            WithLabel("LinkedIn").
            WithSuperTip("Open LinkedIn's website in the default browser.").
            WithImage(LoadBitmap("ExampleRibbon.linkedin.png"), AddressOf GetImage),
            template:=buttonWithCustomIconOne,
            "https://www.linkedin.com")

        Dim buttonWithCustomIconThree As Button = New Button(Sub(bb) bb.
            WithLabel("BandCamp").
            WithSuperTip("Open BandCamp's website in the default browser.").
            WithImage(LoadBitmap("ExampleRibbon.bandcamp.png"), AddressOf GetImage),
            template:=buttonWithCustomIconOne,
            "https://bandcamp.com/")

        Dim buttonWithCustomIconOneSmall As Button = New Button(Sub(bb) bb.
            Normal(),
            template:=buttonWithStockIconOne,
            "https://github.com/")

        Dim buttonWithCustomIconTwoSmall As Button = New Button(Sub(bb) bb.
            Normal(),
            template:=buttonWithStockIconTwo,
            "https://www.linkedin.com")

        Dim buttonWithCustomIconThreeSmall As Button = New Button(Sub(bb) bb.
            Normal(),
            template:=buttonWithStockIconThree,
            "https://bandcamp.com/")

        Dim buttonsWithCustomIcons As Group = New Group(
            Sub(gb) gb.WithLabel("Buttons With Custom Icons"),
            {buttonWithCustomIconOne, buttonWithCustomIconTwo, buttonWithCustomIconThree, Separator.CreateVisible(), buttonWithCustomIconOneSmall, buttonWithCustomIconTwoSmall, buttonWithCustomIconThreeSmall})

        Dim textBox As EditBox = New EditBox(Sub(ebb) ebb.
            WithLabel("Editable Text: ").
            WithScreenTip("Editable Text").
            WithSuperTip("You can edit this text, and the updated text will be displayed in the status bar!").
            AsWideAs("Some Text This Big").
            WithText("Edit Me!", AddressOf GetText, AddressOf OnChange).
            BeforeTextChange(Sub(sender, e) If e.NewText.Contains("  ") Then e.Cancel(), Sub(sender, e) If e.IsCancelled Then DisplayStatusBarMessage("Double spaces are not allowed!")).
            AfterTextChange(Sub(sender, e) DisplayStatusBarMessage($"Text was changed to '{e.NewText}'.")))

        Dim comboBox As ComboBox = New ComboBox(Sub(cbb) cbb.
            WithText("Edit Me!", AddressOf GetText, AddressOf OnChange).
            BeforeTextChange(
                             Sub(sender, e) If Not DirectCast(sender, ComboBox).Any(Function(item) item.Label.Equals(e.NewText, StringComparison.OrdinalIgnoreCase)) Then e.Cancel(),
                             Sub(sender, e) If e.IsCancelled Then DisplayStatusBarMessage($"'{e.NewText}' is not one of the available options.")).
            OnTextChange(Sub(sender, e) DisplayStatusBarMessage($"Text was changed to '{e.NewText}'.")).
            GetItemCountFrom(AddressOf GetItemCount).
            GetItemIdFrom(AddressOf GetItemID).
            GetItemLabelFrom(AddressOf GetItemLabel).
            GetItemSuperTipFrom(AddressOf GetItemSuperTip).
            GetItemScreenTipFrom(AddressOf GetItemScreenTip),
            textBox)

        With comboBox
            .Add("Option One", "The first option.")
            .Add("Option Two", "The second option.")
            .Add("Option Three", "The third option.")
        End With

        Dim textBoxGroup As Group = New Group(Sub(gb) gb.WithLabel("Editable Text"), Box.Vertical(textBox, comboBox))

        Dim one As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.One)).
                                             HideLabel().
                                             Normal().
                                             WithImage(Number.One).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")).
                                             RouteClickTo(AddressOf OnAction),
                                       tag:=Number.One.NumericValue)

        Dim two As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Two)).
                                             HideLabel().
                                             WithImage(Number.Two).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Two.NumericValue)

        Dim three As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Three)).
                                             HideLabel().
                                             WithImage(Number.Three).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Three.NumericValue)

        Dim four As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Four)).
                                             HideLabel().
                                             WithImage(Number.Four).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Four.NumericValue)

        Dim five As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Five)).
                                             HideLabel().
                                             WithImage(Number.Five).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Five.NumericValue)

        Dim six As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Six)).
                                             HideLabel().
                                             WithImage(Number.Six).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Six.NumericValue)

        Dim seven As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Seven)).
                                             HideLabel().
                                             WithImage(Number.Seven).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Seven.NumericValue)

        Dim eight As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Eight)).
                                             HideLabel().
                                             WithImage(Number.Eight).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Eight.NumericValue)

        Dim nine As Button = New Button(Sub(bb) bb.
                                             WithLabel(NameOf(Number.Nine)).
                                             HideLabel().
                                             WithImage(Number.Nine).
                                             OnClick(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!")),
                                       template:=one,
                                       tag:=Number.Nine.NumericValue)

        Dim numbersTopRow As ButtonGroup = New ButtonGroup(items:=ButtonGroupControls.From(one, two, three))

        Dim numbersMiddleRow As ButtonGroup = New ButtonGroup(items:=ButtonGroupControls.From(four, five, six))

        Dim numbersBottomRow As ButtonGroup = New ButtonGroup(items:=ButtonGroupControls.From(seven, eight, nine))

        Dim numberGroup As Group = New Group(Sub(gb) gb.WithLabel("Numbers"), {numbersTopRow, numbersMiddleRow, numbersBottomRow})

        Dim heart As Button = New Button(Sub(bb) bb.
            WithLabel("Heart").
            WithSuperTip("A heart.").
            WithDescription("This suit was invented in 15th century Germany and is a survivor from a large pool of experimental suit signs created to replace the Latin suits.").
            WithImage(Misc.Heart),
            template:=buttonWithStockIconOne)

        Dim spade As Button = New Button(Sub(bb) bb.
            WithLabel("Spade").
            WithSuperTip("A black spade.").
            WithDescription("Spades form one of the four suits of playing cards in the standard French deck. It is a black heart turned upside down with a stalk at its base and symbolises the pike or halberd, two medieval weapons.").
            WithImage(Misc.Spade),
            template:=heart)

        Dim club As Button = New Button(Sub(bb) bb.
            WithLabel("Club").
            WithSuperTip("A black club.").
            WithDescription("Clubs is one of the four suits of playing cards in the standard French deck. It corresponds to the suit of Acorns in a German deck.").
            WithImage(Misc.Club),
            template:=heart)

        Dim diamond As Button = New Button(Sub(bb) bb.
            WithLabel("Diamond").
            WithSuperTip("A red diamond.").
            WithDescription("Diamonds is one of the four suits of playing cards in the standard French deck. It is the only French suit to not have been adapted from the German deck, taking the place of the suit of Bells Bay.").
            WithImage(Misc.Diamond),
            template:=heart)

        Dim cardsMenu As Menu = New Menu(Sub(mb) mb.
            Large().
            WithLargeItems().
            WithLabel("Cards").
            WithSuperTip("Pick a card!").
            WithImage(LoadBitmap("ExampleRibbon.playing-cards-icon.png"), AddressOf GetImage),
            items:=MenuControls.From(heart, spade, club, diamond))

        Dim cardsGroup As Group = New Group(Nothing, {cardsMenu}, template:=cardsMenu)

        Dim desktopFilesDropdownGroup As Group = New DesktopFilesGroup(Me).AsGroup()

        'TODO Finish this
        Dim gallery As Gallery = New Gallery(Sub(gb) gb.
            Large().
            WithImage(Common.Refresh).
            WithLabel("My Custom Gallery").
            WithColumnCount(5).
            WithRowCount(2).
            WithItemDimensions(New Size(32, 32)).
            GetItemCountFrom(AddressOf GetItemCount).
            GetItemImageFrom(AddressOf GetItemImage).
            GetSelectedItemIndexFrom(getSelected:=AddressOf GetSelectedItemIndex, setSelected:=AddressOf OnSelectionChange))

        For i As Integer = 1 To (gallery.Rows * gallery.Columns)
            gallery.Add(New Item(config:=Sub(ib) ib.WithImage(LoadBitmap("ExampleRibbon.bandcamp.png"))))
        Next

        Dim galleryGroup As Group = New Group(config:=Sub(gb) gb.WithLabel("Gallery"), {gallery})

        Dim tab As Tab = New Tab(config:=Sub(tb) tb.
                WithLabel("My Custom Tab").
                InsertAfter(BuiltIn.Excel.TabHome),
                groups:={buttonsWithStockIcons, buttonsWithCustomIcons, textBoxGroup, numberGroup, cardsGroup, desktopFilesDropdownGroup, galleryGroup})

        Return New Controls.Ribbon(Sub(rb) rb.OnLoad(AddressOf OnLoad), tab)
    End Function

    Private Shared Sub OpenWebsiteInDefaultBrowser(webAddress As String)
        Try
            Using process As Process = New Process()
                Process.Start(webAddress)
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private _source As CancellationTokenSource

    Private Async Sub DisplayStatusBarMessage(message As String)
        If _source IsNot Nothing Then
            _source.Cancel()
        End If

        Dim newSource As CancellationTokenSource = New CancellationTokenSource()
        _source = newSource

        Try
            With Excel
                .DisplayStatusBar = True
                .StatusBar = message
            End With

            Await Task.Delay(5000, _source.Token)

            Excel.StatusBar = False
        Catch ex As OperationCanceledException

        Catch ex As Exception

        End Try

        If _source Is newSource Then
            _source = Nothing
        End If
    End Sub

    Private Sub SetContentsOfSelectedCell(contents As Object)
        If Excel.ActiveCell IsNot Nothing Then
            Excel.ActiveCell.Value2 = contents
        End If
    End Sub

End Class