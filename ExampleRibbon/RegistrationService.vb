Imports Microsoft.Win32
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class RegistrationService
	Private ReadOnly info As ManagedOfficeComAddInRegistrationInfo

	Public Sub New(objectToBeRegistered As Object)
		info = New ManagedOfficeComAddInRegistrationInfo(objectToBeRegistered)
	End Sub

	Private NotInheritable Class ManagedOfficeComAddInRegistrationInfo
		Public Sub New(connectionClass As Object)
			With connectionClass.GetType()
				ProgId = .GetCustomAttribute(Of ProgIdAttribute).Value
				Guid = .GetCustomAttribute(Of GuidAttribute).Value

				With .Assembly
					AssemblyName = .FullName
					DllLocation = .Location
				End With
			End With

			DotNetCategoryId = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
		End Sub

		Public Readonly Property ProgId As String

		Public Readonly Property Guid As String

		Public Readonly Property AssemblyName As String

		Public Readonly Property DllLocation As String

		Public ReadOnly Property DotNetCategoryId As String
	End Class

	Public Sub RegisterManagedComAddIn()
		Registry.SetValue($"HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\{info.ProgId}", String.Empty, info.ProgId, RegistryValueKind.String)

		Registry.SetValue($"HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\{info.ProgId}\\CLSID", String.Empty, info.Guid)

		Registry.SetValue($"HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\CLSID\\{info.Guid}", String.Empty, info.ProgId, RegistryValueKind.String)

		Registry.SetValue($"HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\CLSID\\{info.Guid}\\ProgId", String.Empty, info.ProgId, RegistryValueKind.String)

		Using inProcessServerEntries As RegistryKey = Registry.LocalMachine.CreateSubKey($"SOFTWARE\\Classes\\CLSID\\{info.Guid}\\inProcServer32")
			With inProcessServerEntries
				.SetValue(String.Empty, "mscoree.dll")
				.SetValue("ThreadingModel", "Both")
				.SetValue("Class", info.ProgId)
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using inProcessServerEntries As RegistryKey = Registry.LocalMachine.CreateSubKey($"SOFTWARE\\Classes\\CLSID\\{info.Guid}\\inProcServer32\\1.0.0.0")
			With inProcessServerEntries
				.SetValue("Class", info.ProgId)
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using implementedCategoryEntry As RegistryKey = Registry.LocalMachine.CreateSubKey($"SOFTWARE\\Classes\\CLSID\\{info.Guid}\\Implemented Categories\\{info.DotNetCategoryId}")

		End Using
	End Sub

	Public Sub UnRegisterManagedComAddIn()
		Using localMachine As RegistryKey = Registry.LocalMachine
			With localMachine
				.DeleteSubKeyTree($"SOFTWARE\\Classes\\{info.ProgId}")
				.DeleteSubKeyTree($"SOFTWARE\\Classes\\CLSID\\{info.Guid}")
			End With
		End Using
	End Sub
End Class
