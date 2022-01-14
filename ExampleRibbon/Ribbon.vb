Imports System.ComponentModel
Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Extensibility
Imports Microsoft.Office.Core
Imports RibbonFactory
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports Excel = Microsoft.Office.Interop.Excel

<ComVisible(True)>
<Guid("C2C29BAF-8F1B-46EF-A071-8A286423F4C4")>
<ProgId("ExampleRibbon.Ribbon")>
Public Class Ribbon
	Inherits RibbonFactory.Utilities.StockRibbonBase
	Implements IDTExtensibility2
	Implements IRibbonExtensibility

	Public Sub New()
		MyBase.New()
	End Sub

	Private ReadOnly Property Excel As Excel.Application
		Get
			Return DirectCast(HostApp, Excel.Application)
		End Get
	End Property

	Protected Overrides Function BuildRibbon() As Containers.Ribbon
		Dim buttonWithStockIconOne As Button = New ButtonBuilder().
			Large().
			WithLabel("Happy Button").
			WithSuperTip("Oh, to be so happy again!").
			WithImage(Enums.ImageMSO.Common.HappyFace).
			ThatDoes(Sub(b) DisplayStatusBarMessage($"You clicked '{b.Label}'!"), AddressOf OnAction).
			Build()

		Dim buttonWithStockIconTwo As Button = New ButtonBuilder(buttonWithStockIconOne).
			WithLabel("Sad Button").
			WithSuperTip("#Sad").
			WithImage(Enums.ImageMSO.Common.SadFace).
			Build()

		Dim buttonWithStockIconThree As Button = New ButtonBuilder(buttonWithStockIconOne).
			WithLabel("Money Button").
			WithSuperTip("Make that money!").
			WithImage(Enums.ImageMSO.Common.DollarSign).
			Build()

		Dim buttonWithStockIconsSeparator As Separator = New SeparatorBuilder().
			Visible().
			Build()

		Dim buttonWithStockIconOneSmall As Button = New ButtonBuilder(buttonWithStockIconOne).
			Normal().
			Build()

		Dim buttonWithStockIconTwoSmall As Button = New ButtonBuilder(buttonWithStockIconTwo).
			Normal().
			Build()

		Dim buttonWithStockIconThreeSmall As Button = New ButtonBuilder(buttonWithStockIconThree).
			Normal().
			Build()

		Dim buttonsWithStockIcons As Group = New GroupBuilder().
			WithControls(buttonWithStockIconOne, buttonWithStockIconTwo, buttonWithStockIconThree, buttonWithStockIconsSeparator, buttonWithStockIconOneSmall, buttonWithStockIconTwoSmall, buttonWithStockIconThreeSmall).
			WithLabel("Buttons With Stock Icons").
			Build()

		Dim buttonWithCustomIconOne As Button = New ButtonBuilder().
			Large().
			WithLabel("GitHub").
			WithSuperTip("Open GitHub's website in the default browser.").
			WithImage(LoadBitmap("ExampleRibbon.github.png"), AddressOf GetImage).
			ThatDoes(Sub(b) OpenWebsiteInDefaultBrowser(b.Tag.ToString()), AddressOf OnAction).
			Build(tag:="https://github.com/")

		Dim buttonWithCustomIconTwo As Button = New ButtonBuilder(buttonWithCustomIconOne).
			WithLabel("LinkedIn").
			WithSuperTip("Open LinkedIn's website in the default browser.").
			WithImage(LoadBitmap("ExampleRibbon.linkedin.png"), AddressOf GetImage).
			Build(tag:="https://www.linkedin.com")

		Dim buttonWithCustomIconThree As Button = New ButtonBuilder(buttonWithCustomIconOne).
			WithLabel("BandCamp").
			WithSuperTip("Open BandCamp's website in the default browser.").
			WithImage(LoadBitmap("ExampleRibbon.bandcamp.png"), AddressOf GetImage).
			Build(tag:="https://bandcamp.com/")

		Dim buttonWithCustomIconsSeparator As Separator = New SeparatorBuilder().
			Visible().
			Build()

		Dim buttonWithCustomIconOneSmall As Button = New ButtonBuilder(buttonWithCustomIconOne).
			Normal().
			Build(tag:="https://github.com/")

		Dim buttonWithCustomIconTwoSmall As Button = New ButtonBuilder(buttonWithCustomIconTwo).
			Normal().
			Build(tag:="https://www.linkedin.com")

		Dim buttonWithCustomIconThreeSmall As Button = New ButtonBuilder(buttonWithCustomIconThree).
			Normal().
			Build(tag:="https://bandcamp.com/")

		Dim buttonsWithCustomIcons As Group = New GroupBuilder().
			WithControls(buttonWithCustomIconOne, buttonWithCustomIconTwo, buttonWithCustomIconThree, buttonWithCustomIconsSeparator, buttonWithCustomIconOneSmall, buttonWithCustomIconTwoSmall, buttonWithCustomIconThreeSmall).
			WithLabel("Buttons With Custom Icons").
			Build()

		Dim textBox As EditBox = New EditBoxBuilder().
			WithLabel("Editable Text: ").
			WithScreenTip("Editable Text").
			WithSuperTip("You can edit this text, and the updated text will be displayed in the status bar!").
			AsWideAs("Some Text This Big").
			WithText("Edit Me!", AddressOf GetText).
			ThatDoes(Sub(t) Return, AddressOf OnChange).
			WithTextValidationRule(Function(text) Not text.Contains("  "), "Double-Spaces Aren't Allowed!").
			Build()

		AddHandler textBox.TextChanged, AddressOf TextChangeEventHandler

		Dim textBoxGroup As Group = New GroupBuilder().WithLabel("Editable Text").WithControl(textBox).Build()

		Dim one As Button = New ButtonBuilder().
			WithScreenTip("One").
			WithImage(Enums.ImageMSO.Number.One).
			ThatDoes(Sub(b) SetContentsOfSelectedCell(b.Tag), AddressOf OnAction).
			Build(1)

		Dim two As Button = New ButtonBuilder(one).
			WithScreenTip("Two").
			WithImage(Enums.ImageMSO.Number.Two).
			Build(2)

		Dim three As Button = New ButtonBuilder(one).
			WithScreenTip("Three").
			WithImage(Enums.ImageMSO.Number.Three).
			Build(3)

		Dim four As Button = New ButtonBuilder(one).
			WithScreenTip("Four").
			WithImage(Enums.ImageMSO.Number.Four).
			Build(4)

		Dim five As Button = New ButtonBuilder(one).
			WithScreenTip("Five").
			WithImage(Enums.ImageMSO.Number.Five).
			Build(5)

		Dim six As Button = New ButtonBuilder(one).
			WithScreenTip("Six").
			WithImage(Enums.ImageMSO.Number.Six).
			Build(6)

		Dim seven As Button = New ButtonBuilder(one).
			WithScreenTip("Seven").
			WithImage(Enums.ImageMSO.Number.Seven).
			Build(7)

		Dim eight As Button = New ButtonBuilder(one).
			WithScreenTip("Eight").
			WithImage(Enums.ImageMSO.Number.Eight).
			Build(8)

		Dim nine As Button = New ButtonBuilder(one).
			WithScreenTip("Nine").
			WithImage(Enums.ImageMSO.Number.Nine).
			Build(9)

		Dim numbersTopRow As ButtonGroup = New ButtonGroupBuilder().
			WithControls(one, two, three).
			Build()

		Dim numbersMiddleRow As ButtonGroup = New ButtonGroupBuilder().
			WithControls(four, five, six).
			Build()

		Dim numbersBottomRow As ButtonGroup = New ButtonGroupBuilder().
			WithControls(seven, eight, nine).
			Build()

		Dim numberGroup As Group = New GroupBuilder().
			WithControls(numbersTopRow, numbersMiddleRow, numbersBottomRow).
			WithLabel("Numbers").
			Build()

		Dim heart As Button = New ButtonBuilder(buttonWithStockIconOne).
			WithLabel("Heart").
			WithSuperTip("A heart.").
			WithDescription("This suit was invented in 15th century Germany and is a survivor from a large pool of experimental suit signs created to replace the Latin suits.").
			WithImage(Enums.ImageMSO.Misc.Heart).
			Build()

		Dim spade As Button = New ButtonBuilder(heart).
			WithLabel("Spade").
			WithSuperTip("A black spade.").
			WithDescription("Spades form one of the four suits of playing cards in the standard French deck. It is a black heart turned upside down with a stalk at its base and symbolises the pike or halberd, two medieval weapons.").
			WithImage(Enums.ImageMSO.Misc.Spade).
			Build()

		Dim club As Button = New ButtonBuilder(heart).
			WithLabel("Club").
			WithSuperTip("A black club.").
			WithDescription("Clubs is one of the four suits of playing cards in the standard French deck. It corresponds to the suit of Acorns in a German deck.").
			WithImage(Enums.ImageMSO.Misc.Club).
			Build()

		Dim diamond As Button = New ButtonBuilder(heart).
			WithLabel("Diamond").
			WithSuperTip("A red diamond.").
			WithDescription("Diamonds is one of the four suits of playing cards in the standard French deck. It is the only French suit to not have been adapted from the German deck, taking the place of the suit of Bells Bay.").
			WithImage(Enums.ImageMSO.Misc.Diamond).
			Build()

		Dim cardsMenu As Menu = New MenuBuilder().
			Large().
			WithLargeItems().
			WithLabel("Cards").
			WithSuperTip("Pick a card!").
			WithImage(LoadBitmap("ExampleRibbon.playing-cards-icon.png"), AddressOf GetImage).
			WithControls(heart, spade, club, diamond).
			Build()

		Dim cardsGroup As Group = New GroupBuilder(cardsMenu).
			WithControl(cardsMenu).
			Build()

		Dim desktopFilesDropdownGroup As Group = New DesktopFilesGroup(Me).AsGroup()

		Dim tab As Tab = New TabBuilder().
				WithGroups(buttonsWithStockIcons, buttonsWithCustomIcons, textBoxGroup, numberGroup, cardsGroup, desktopFilesDropdownGroup).
				WithLabel("My Custom Tab").
				InsertAfterMSO(Enums.MSO.Excel.TabHome).
				Build()

		Return New RibbonBuilder().
			OnLoad(AddressOf OnLoad).
			WithTab(tab).
			Build()
	End Function

	Private Shared Function LoadBitmap(path As String) As Bitmap
		Return New Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream(path))
	End Function

	Private Shared Sub OpenWebsiteInDefaultBrowser(webAddress As String)
		Try
			Using process As Process = New Process()
				Process.Start(webAddress)
			End Using
		Catch ex As Exception

		End Try
	End Sub

	Private Sub TextChangeEventHandler(sender As Object, e As EditBox.TextChangedEventArgs)
		DisplayStatusBarMessage(If(e.IsSuccess, $"Text was changed to '{e.Text}'.", e.ErrorMessage))
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