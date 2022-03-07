Imports Microsoft.Win32
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class RegistrationService
	Private ReadOnly info As ManagedOfficeComAddInRegistrationInfo

	Public Sub New(type As Type, friendlyName As String, description As String, targetApplication As String, loadBehavior As Byte)
		info = New ManagedOfficeComAddInRegistrationInfo(type, friendlyName, description, loadBehavior, targetApplication)
	End Sub

	Private NotInheritable Class ManagedOfficeComAddInRegistrationInfo
		Public Sub New(type As Type, friendlyName As String, description As String, loadBehavior As Byte, targetApplication As String)
			With type
				ProgId = .GetCustomAttribute(Of ProgIdAttribute).Value
				Guid = "{" & .GetCustomAttribute(Of GuidAttribute).Value & "}"

				With .Assembly
					AssemblyName = .FullName
					DllLocation = .Location
				End With
			End With

			DotNetCategoryId = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
			Me.FriendlyName = friendlyName
			Me.Description = description
			Me.LoadBehavior = loadBehavior
			Me.TargetApplication = targetApplication
		End Sub

		Public Readonly Property ProgId As String

		Public Readonly Property Guid As String

		Public Readonly Property AssemblyName As String

		Public Readonly Property DllLocation As String

		Public ReadOnly Property DotNetCategoryId As String

		Public ReadOnly Property FriendlyName As String

		Public Readonly Property Description As String

		Public ReadOnly Property LoadBehavior As Byte

		Public Readonly Property TargetApplication As String
	End Class

	Public Sub RegisterManagedComAddIn()

		Registry.SetValue($"HKEY_CURRENT_USER\SOFTWARE\Classes\{info.ProgId}", String.Empty, info.ProgId, RegistryValueKind.String)

		Registry.SetValue($"HKEY_CURRENT_USER\SOFTWARE\Classes\{info.ProgId}\CLSID", String.Empty, info.Guid)

		Registry.SetValue($"HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID\{info.Guid}", String.Empty, info.ProgId, RegistryValueKind.String)

		Registry.SetValue($"HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID\{info.Guid}\ProgId", String.Empty, info.ProgId, RegistryValueKind.String)

		Registry.SetValue($"HKEY_CURRENT_USER\SOFTWARE\Classes\Wow6432Node\CLSID\{info.Guid}\ProgId", String.Empty, info.ProgId, RegistryValueKind.String)

		Using inProcessServerEntries As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Classes\CLSID\{info.Guid}\inProcServer32")
			With inProcessServerEntries
				.SetValue(String.Empty, "mscoree.dll")
				.SetValue("ThreadingModel", "Both")
				.SetValue("Class", info.ProgId)
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using inProcessServerEntries As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Classes\CLSID\{info.Guid}\inProcServer32\1.0.0.0")
			With inProcessServerEntries
				.SetValue("Class", info.ProgId)
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using implementedCategoryEntry As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Classes\CLSID\{info.Guid}\Implemented Categories\{info.DotNetCategoryId}")

		End Using

		Using implementedCategoryEntry As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Classes\Wow6432Node\CLSID\{info.Guid}\Implemented Categories\{info.DotNetCategoryId}")

		End Using

		Using classRecord As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Classes\Record\{289E9AF1-4973-11D1-AE81-00A0C90F26F4}\1.0.0.0")
			With classRecord
				.SetValue("Class", "Extensibility.ext_ConnectMode")
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using classRecord As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Classes\Record\{289E9AF2-4973-11D1-AE81-00A0C90F26F4}\1.0.0.0")
			With classRecord
				.SetValue("Class", "Extensibility.ext_DisconnectMode")
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using inProcessServerEntries As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Classes\Wow6432Node\CLSID\{info.Guid}\inProcServer32")
			With inProcessServerEntries
				.SetValue(String.Empty, "mscoree.dll")
				.SetValue("ThreadingModel", "Both")
				.SetValue("Class", info.ProgId)
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using inProcessServerEntries As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Classes\Wow6432Node\CLSID\{info.Guid}\inProcServer32\1.0.0.0")
			With inProcessServerEntries
				.SetValue("Class", info.ProgId)
				.SetValue("Assembly", info.AssemblyName)
				.SetValue("RuntimeVersion", "v4.0.30319")
				.SetValue("CodeBase", $"file:///{info.DllLocation}")
			End With
		End Using

		Using applicationSpecific As RegistryKey = Registry.CurrentUser.CreateSubKey($"SOFTWARE\Microsoft\Office\{info.TargetApplication}\Addins\{info.ProgId}")
			With applicationSpecific
				.SetValue("FriendlyName", info.FriendlyName, RegistryValueKind.String)
				.SetValue("Description", info.Description, RegistryValueKind.String)
				.SetValue("LoadBehavior", info.LoadBehavior, RegistryValueKind.DWord)
			End With
		End Using
	End Sub

	Public Sub UnRegisterManagedComAddIn()
		Using currentUser As RegistryKey = Registry.CurrentUser
			With currentUser
				If .OpenSubKey($"SOFTWARE\Classes\{info.ProgId}") IsNot Nothing Then
					.DeleteSubKeyTree($"SOFTWARE\Classes\{info.ProgId}")
				End If
				
				If .OpenSubKey($"SOFTWARE\Classes\CLSID\{info.Guid}") IsNot Nothing Then
					.DeleteSubKeyTree($"SOFTWARE\Classes\CLSID\{info.Guid}")
				End If

				If .OpenSubKey($"SOFTWARE\Classes\Wow6432Node\CLSID\{info.Guid}") IsNot Nothing Then
					.DeleteSubKeyTree($"SOFTWARE\Classes\Wow6432Node\CLSID\{info.Guid}")
				End If
				
				If .OpenSubKey($"SOFTWARE\Microsoft\Office\{info.TargetApplication}\Addins\{info.ProgId}") IsNot Nothing Then
					.DeleteSubKeyTree($"SOFTWARE\Microsoft\Office\{info.TargetApplication}\Addins\{info.ProgId}")
				End If
			End With
		End Using
	End Sub
End Class
