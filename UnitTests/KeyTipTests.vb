Imports RibbonFactory

<TestClass()>
Public Class KeyTipTests

	<TestMethod()>
	Public Sub NumericConversionToA()
		Dim tip As KeyTip = 1

		Assert.AreEqual("A", tip.ToString())
	End Sub

	<TestMethod()>
	Public Sub NumericConversionToA1()
		Dim tip As KeyTip = 62

		Assert.AreEqual("A1", tip.ToString())
	End Sub

	<TestMethod()>
	Public Sub StringConversionToA2()
		Dim tip As KeyTip = "A2"

		Assert.AreEqual("A2", tip.ToString())
	End Sub

	<TestMethod()>
	Public Sub StringConversionToRS()
		Dim tip As KeyTip = "RS"

		Assert.AreEqual("RS", tip.ToString())
	End Sub

	<TestMethod()>
	Public Sub ThrowsInvalidFormatExceptionWhenAnyCharacterOtherThanTheLastCharacterIsANumber()
		Dim tip As KeyTip

		Assert.ThrowsException(Of FormatException)(Sub() tip = "12")
	End Sub

	<TestMethod()>
	Public Sub GetNextAvailableWorks()
		Debug.WriteLine($"The next available KeyTip is: '{KeyTip.NextAvailable()}'")
	End Sub

End Class