Imports System.Runtime.CompilerServices

Friend Module RandomExtensions

	<Extension()>
	Public Function NextBoolean(random As Random) As Boolean
		Return random.Next Mod 2 = 0
	End Function

End Module
