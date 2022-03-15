Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls

Public Interface IButtonBuilder
	Inherits IID(Of IButtonBuilder)
	Inherits IEnabled(Of IButtonBuilder)
	Inherits IVisible(Of IButtonBuilder)
	Inherits ILabelScreenTipSuperTip(Of IButtonBuilder)
	Inherits IShowLabel(Of IButtonBuilder)
	Inherits IShowImage(Of IButtonBuilder)
	Inherits ISize(Of IButtonBuilder)
	Inherits IImage(Of IButtonBuilder)
	Inherits IDescription(Of IButtonBuilder)
	Inherits IKeyTip(Of IButtonBuilder)
	Inherits IInsert(Of IButtonBuilder)

	Function BeforeClick(callBack As OnAction, ParamArray actions() As Action(Of Button, Button.BeforeClickEventArgs)) As IButtonBuilder
	Function OnClick(callBack As OnAction, ParamArray actions() As Action(Of Button)) As IButtonBuilder
End Interface
