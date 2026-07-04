Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports RibbonX
Imports RibbonX.ComTypes.Extensibility
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.Controls
Imports RibbonX.Images.BuiltIn
Imports Excel = Microsoft.Office.Interop.Excel


<ComVisible(True)>
<Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4")>
<ProgId("ExampleRibbon.Ribbon")>
Public Class Ribbon
    Inherits CustomRibbonBase
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

    Protected Overrides Function BuildRibbon() As IRibbon
        Dim buttonWithStockIconOne As IButton = RibbonXBuilder.Button(
                Sub(bb) bb.
                Large().
                WithLabel("Happy Button").
                WithSuperTip("Oh, to be so happy again!").
                WithImage(Common.HappyFace).
                OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(btn) DisplayStatusBarMessage($"You clicked '{btn.Label}'!"))))

        Dim buttonWithStockIconTwo As IButton = RibbonXBuilder.Button(Sub(bb) bb.
                FromTemplate(buttonWithStockIconOne).
                WithLabel("Sad Button").
                WithSuperTip("#Sad").
                WithImage(Common.SadFace).
                OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(btn) DisplayStatusBarMessage($"You clicked '{btn.Label}'!"))))

        Dim buttonWithStockIconThree As IButton = RibbonXBuilder.Button(Sub(bb) bb.
                FromTemplate(buttonWithStockIconOne).
                WithLabel("Money Button").
                WithSuperTip("Make that money!").
                WithImage(Common.DollarSign).
                OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(btn) DisplayStatusBarMessage($"You clicked '{btn.Label}'!"))))


        Dim buttonWithStockIconOneSmall As IButton = RibbonXBuilder.Button(Sub(bb) bb.FromTemplate(buttonWithStockIconOne).Normal())

        Dim buttonWithStockIconTwoSmall As IButton = RibbonXBuilder.Button(Sub(bb) bb.FromTemplate(buttonWithStockIconTwo).Normal())

        Dim buttonWithStockIconThreeSmall As IButton = RibbonXBuilder.Button(Sub(bb) bb.FromTemplate(buttonWithStockIconThree).Normal())

        Dim buttonsWithStockIcons As IGroup = RibbonXBuilder.Group(Sub(gb) gb.
            WithLabel("Buttons With Stock Icons").
            WithControls(buttonWithStockIconOne, buttonWithStockIconTwo, buttonWithStockIconThree, RibbonXBuilder.Separator(), buttonWithStockIconOneSmall, buttonWithStockIconTwoSmall, buttonWithStockIconThreeSmall))

        Dim buttonWithCustomIconOne As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            Large().
            WithLabel("GitHub").
            WithSuperTip("Open GitHub's website in the default browser.").
            WithImage(LoadBitmap("ExampleRibbon.github.png"), AddressOf GetImage).
            WithTag("https://github.com/").
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(btn) OpenWebsiteInDefaultBrowser(btn.Tag.ToString()), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim buttonWithCustomIconTwo As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(buttonWithCustomIconOne).
            WithLabel("LinkedIn").
            WithSuperTip("Open LinkedIn's website in the default browser.").
            WithImage(LoadBitmap("ExampleRibbon.linkedin.png"), AddressOf GetImage).
            WithTag("https://www.linkedin.com").
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(btn) OpenWebsiteInDefaultBrowser(btn.Tag.ToString()), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim buttonWithCustomIconThree As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(buttonWithCustomIconOne).
            WithLabel("BandCamp").
            WithSuperTip("Open BandCamp's website in the default browser.").
            WithImage(LoadBitmap("ExampleRibbon.bandcamp.png"), AddressOf GetImage).
            WithTag("https://bandcamp.com/"))

        Dim buttonWithCustomIconOneSmall As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(buttonWithStockIconOne).
            Normal().
            WithTag("https://github.com/"))

        Dim buttonWithCustomIconTwoSmall As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(buttonWithStockIconTwo).
            Normal().
            WithTag("https://www.linkedin.com"))

        Dim buttonWithCustomIconThreeSmall As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(buttonWithStockIconThree).
            Normal().
            WithTag("https://bandcamp.com/"))

        Dim buttonsWithCustomIcons As IGroup = RibbonXBuilder.Group(
            Sub(gb) gb.WithLabel("Buttons With Custom Icons").
            WithControls(buttonWithCustomIconOne, buttonWithCustomIconTwo, buttonWithCustomIconThree, RibbonXBuilder.Separator(), buttonWithCustomIconOneSmall, buttonWithCustomIconTwoSmall, buttonWithCustomIconThreeSmall))

        Dim textBox As IEditBox = RibbonXBuilder.EditBox(Sub(ebb) ebb.
            WithLabel("Editable Text: ").
            WithScreenTip("Editable Text").
            WithSuperTip("You can edit this text, and the updated text will be displayed in the status bar!").
            AsWideAs("Some Text This Big").
            WithText("Edit Me!", AddressOf GetText, AddressOf OnChange,
                Sub(action)
                    action.ButFirst(Function(newText As String) Not newText.Contains("  "))
                    action.Do(Sub(newText As String) DisplayStatusBarMessage($"Text was changed to '{newText}'."))
                End Sub))

        Dim comboBox As IComboBox = RibbonXBuilder.ComboBox(Sub(cbb) cbb.
            WithText("Edit Me!", AddressOf GetText, AddressOf OnChange,
                Sub(action)
                    action.ButFirst(Function(cb As IComboBox, newText As String) cb.Any(Function(item) item.Label.Equals(newText, StringComparison.OrdinalIgnoreCase)))
                    action.Do(Sub(newText As String) DisplayStatusBarMessage($"Text was changed to '{newText}'."))
                End Sub).
            GetItemCountFrom(AddressOf GetItemCount).
            GetItemIdFrom(AddressOf GetItemID).
            GetItemLabelFrom(AddressOf GetItemLabel).
            GetItemSuperTipFrom(AddressOf GetItemSuperTip).
            GetItemScreenTipFrom(AddressOf GetItemScreenTip))

        With comboBox
            .Add(RibbonXBuilder.Item(Sub(ib) ib.WithLabel("Option One").WithScreenTip("The first option.")))
            .Add(RibbonXBuilder.Item(Sub(ib) ib.WithLabel("Option Two").WithScreenTip("The second option.")))
            .Add(RibbonXBuilder.Item(Sub(ib) ib.WithLabel("Option Three").WithScreenTip("The third option.")))
        End With

        Dim textBoxGroup As IGroup = RibbonXBuilder.Group(Sub(gb) gb.WithLabel("Editable Text").WithControls(RibbonXBuilder.Box(Sub(bb) bb.Vertical().WithControls(textBox, comboBox))))

        Dim one As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            WithLabel(NameOf(Number.One)).
            HideLabel().
            Normal().
            WithImage(Number.One).
            WithTag(Number.One.NumericValue).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim two As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Two.NumericValue).
            WithLabel(NameOf(Number.Two)).
            HideLabel().
            WithImage(Number.Two).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim three As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Three.NumericValue).
            WithLabel(NameOf(Number.Three)).
            HideLabel().
            WithImage(Number.Three).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim four As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Four.NumericValue).
            WithLabel(NameOf(Number.Four)).
            HideLabel().
            WithImage(Number.Four).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim five As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Five.NumericValue).
            WithLabel(NameOf(Number.Five)).
            HideLabel().
            WithImage(Number.Five).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim six As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Six.NumericValue).
            WithLabel(NameOf(Number.Six)).
            HideLabel().
            WithImage(Number.Six).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim seven As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Seven.NumericValue).
            WithLabel(NameOf(Number.Seven)).
            HideLabel().
            WithImage(Number.Seven).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim eight As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Eight.NumericValue).
            WithLabel(NameOf(Number.Eight)).
            HideLabel().
            WithImage(Number.Eight).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim nine As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(one).
            WithTag(Number.Nine.NumericValue).
            WithLabel(NameOf(Number.Nine)).
            HideLabel().
            WithImage(Number.Nine).
            OnClick(AddressOf OnAction, Sub(click) click.Do(Sub(b) SetContentsOfSelectedCell(b.Tag), Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"))))

        Dim numbersTopRow As IButtonGroup = RibbonXBuilder.ButtonGroup(Sub(bg) bg.WithControls(one, two, three))

        Dim numbersMiddleRow As IButtonGroup = RibbonXBuilder.ButtonGroup(Sub(bg) bg.WithControls(four, five, six))

        Dim numbersBottomRow As IButtonGroup = RibbonXBuilder.ButtonGroup(Sub(bg) bg.WithControls(seven, eight, nine))

        Dim numberGroup As IGroup = RibbonXBuilder.Group(Sub(gb) gb.WithLabel("Numbers").WithControls(numbersTopRow, numbersMiddleRow, numbersBottomRow))

        Dim heart As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(buttonWithStockIconOne).
            WithLabel("Heart").
            WithSuperTip("A heart.").
            WithDescription("This suit was invented in 15th century Germany and is a survivor from a large pool of experimental suit signs created to replace the Latin suits.").
            WithImage(Misc.Heart))

        Dim spade As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(heart).
            WithLabel("Spade").
            WithSuperTip("A black spade.").
            WithDescription("Spades form one of the four suits of playing cards in the standard French deck. It is a black heart turned upside down with a stalk at its base and symbolises the pike or halberd, two medieval weapons.").
            WithImage(Misc.Spade))

        Dim club As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(heart).
            WithLabel("Club").
            WithSuperTip("A black club.").
            WithDescription("Clubs is one of the four suits of playing cards in the standard French deck. It corresponds to the suit of Acorns in a German deck.").
            WithImage(Misc.Club))

        Dim diamond As IButton = RibbonXBuilder.Button(Sub(bb) bb.
            FromTemplate(heart).
            WithLabel("Diamond").
            WithSuperTip("A red diamond.").
            WithDescription("Diamonds is one of the four suits of playing cards in the standard French deck. It is the only French suit to not have been adapted from the German deck, taking the place of the suit of Bells Bay.").
            WithImage(Misc.Diamond))

        Dim cardsMenu As IMenu = RibbonXBuilder.Menu(Sub(mb) mb.
            Large().
            WithLargeItems().
            WithLabel("Cards").
            WithSuperTip("Pick a card!").
            WithImage(LoadBitmap("ExampleRibbon.playing-cards-icon.png"), AddressOf GetImage).
            WithControls(heart, spade, club, diamond))

        Dim cardsGroup As IGroup = RibbonXBuilder.Group(Sub(b) b.FromTemplate(cardsMenu).WithControls(cardsMenu))

        Dim desktopFilesDropdownGroup As IGroup = New DesktopFilesGroup(Me).AsGroup()

        'TODO Finish this
        Dim gallery As IGallery = RibbonXBuilder.Gallery(Sub(gb) gb.
            Large().
            WithImage(Common.Refresh).
            WithLabel("My Custom Gallery").
            WithColumnCount(5).
            WithRowCount(2).
            WithItemDimensions(New Size(32, 32)).
            GetItemCountFrom(AddressOf GetItemCount).
            GetItemImageFrom(AddressOf GetItemImage).
            GetSelectedItemIndexFrom(AddressOf GetSelectedItemIndex, AddressOf OnSelectionChange))

        For i As Integer = 1 To (gallery.Rows * gallery.Columns)
            gallery.Add(RibbonXBuilder.Item(Sub(ib) ib.WithImage(LoadBitmap("ExampleRibbon.bandcamp.png"))))
        Next

        Dim galleryGroup As IGroup = RibbonXBuilder.Group(Sub(gb) gb.WithLabel("Gallery").WithControls(gallery))

        Dim tab As ITab = RibbonXBuilder.Tab(Sub(tb) tb.
                WithLabel("My Custom Tab").
                InsertAfter(BuiltIn.Excel.TabHome).
                WithGroups(buttonsWithStockIcons, buttonsWithCustomIcons, textBoxGroup, numberGroup, cardsGroup, desktopFilesDropdownGroup, galleryGroup))

        Return RibbonXBuilder.Ribbon(Sub(rb) rb.OnLoad(AddressOf OnLoad).WithTabs(tab))
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
        _source?.Cancel()

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
