Namespace BuilderInterfaces

	Public Interface IShowInRibbon(Of Out T)

		Function ShowInRibbon() As T

		Function DoNotShowInRibbon() As T

	End Interface

End Namespace

