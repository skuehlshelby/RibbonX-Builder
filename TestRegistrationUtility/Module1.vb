Module Module1

	Sub Main()
		Dim ribbon As ExampleRibbon.Ribbon = New ExampleRibbon.Ribbon()

		ribbon.RemoveRegistryEntries()

		ribbon.CreateRegistryEntries()
	End Sub

End Module
