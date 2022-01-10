Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports RibbonFactory.ControlInterfaces

Namespace Utilities.Testing

	''' <summary>
	''' A test class for simulating requests from an MS Office host application.
	''' </summary>
	Public Class OfficeHostAppilcation
		Implements IRibbonUI

		Private ReadOnly _controls As IDictionary(Of String, RibbonControl)

		Public Sub New(ribbonExtensibility As IRibbonExtensibility)
			Dim temp As String = ribbonExtensibility.GetCustomUI(String.Empty)
			Dim ribbonX As XDocument = XDocument.Parse(temp)

			_controls = ribbonX.
				Root.
				Descendants().
				Where(Function(descendant) descendant.Attribute("id") IsNot Nothing).
				Select(Function(element) New RibbonControl(element, ribbonExtensibility)).
				ToDictionary(Function(control) control.Id)
		End Sub

		''' <summary>
		''' Simulates a click on a ribbon element, as performed by a user in Excel.
		''' </summary>
		''' <param name="element"></param>
		Public Sub Click(element As RibbonElement)
			_controls(element.ID).Click()
		End Sub

		Public Function IsDataSynchronized(element As RibbonElement) As Boolean
			Return _controls(element.ID).CallbackReturnValuesMatchRibbonElementPropertyValues(element)
		End Function

#Region "IRibbonUI Members"

		Public Sub Invalidate() Implements IRibbonUI.Invalidate
			For Each control As RibbonControl In _controls.Values
				control.QueryControlProperties()
			Next
		End Sub

		Public Sub InvalidateControl(controlID As String) Implements IRibbonUI.InvalidateControl
			_controls(controlID).QueryControlProperties()
		End Sub

		Public Sub InvalidateControlMso(ControlID As String) Implements IRibbonUI.InvalidateControlMso
			'Do Nothing
		End Sub

		Public Sub ActivateTab(ControlID As String) Implements IRibbonUI.ActivateTab
			'Do Nothing
		End Sub

		Public Sub ActivateTabMso(ControlID As String) Implements IRibbonUI.ActivateTabMso
			'Do Nothing
		End Sub

		Public Sub ActivateTabQ(ControlID As String, [Namespace] As String) Implements IRibbonUI.ActivateTabQ
			'Do Nothing
		End Sub

#End Region

		Private Class RibbonControl
			Implements IRibbonControl

			Private ReadOnly element As XElement
			Private ReadOnly extensibility As IRibbonExtensibility

			Public Sub New(element As XElement, ribbonExtensibility As IRibbonExtensibility)
				Me.element = element
				extensibility = ribbonExtensibility
			End Sub

			Public ReadOnly Property Id As String Implements IRibbonControl.Id
				Get
					Return If(element.Attribute("id")?.Value, String.Empty)
				End Get
			End Property

			Private ReadOnly Property Context As Object Implements IRibbonControl.Context
				Get
					Return Nothing
				End Get
			End Property

			Private ReadOnly Property Tag As String Implements IRibbonControl.Tag
				Get
					Return element.Attribute("tag")?.Value
				End Get
			End Property

			Public Sub QueryControlProperties()
				For Each attribute As XAttribute In element.Attributes()
					Dim callback As String = Nothing

					If IsCallback(attribute, Nothing, callback) Then
						InvokeMethod(callback)
					End If
				Next
			End Sub

			Public Sub Click()
				Dim action As XAttribute = element.Attribute("onAction")
				InvokeMethod(action.Value)
			End Sub

			Public Sub Type(text As String)
				Dim action As XAttribute = element.Attribute("onChange")
				InvokeMethod(action.Value, text)
			End Sub

			Public Function CallbackReturnValuesMatchRibbonElementPropertyValues(control As RibbonElement) As Boolean
				For Each attribute As XAttribute In element.Attributes()
					Dim name As String = Nothing
					Dim callback As String = Nothing

					'TODO Make sure everything is represented here
					If IsCallback(attribute, name, callback) Then
						Select Case name
							Case "getLabel"
								If Not InvokeMethod(callback).Equals(DirectCast(control, ILabel).Label) Then
									Return False
								End If
							Case "getShowLabel"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IShowLabel).ShowLabel) Then
									Return False
								End If
							Case "getScreentip"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IScreenTip).ScreenTip) Then
									Return False
								End If
							Case "getSupertip"
								If Not InvokeMethod(callback).Equals(DirectCast(control, ISuperTip).SuperTip) Then
									Return False
								End If
							Case "getDescription"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IDescription).Description) Then
									Return False
								End If
							Case "getKeytip"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IKeyTip).KeyTip) Then
									Return False
								End If
							Case "getEnabled"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IEnabled).Enabled) Then
									Return False
								End If
							Case "getVisible"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IVisible).Visible) Then
									Return False
								End If
							Case "getImage"
								If Not InvokeMethod(callback).Equals(DirectCast(control, IImage).Image) Then
									Return False
								End If
							Case "getSize"
								If Not InvokeMethod(callback).Equals(DirectCast(control, ISize).Size) Then
									Return False
								End If
						End Select
					End If
				Next

				Return True
			End Function

			Private Function IsCallback(attribute As XAttribute, <Out> ByRef attributeName As String, <Out> ByRef callback As String) As Boolean
				If attribute.Name.LocalName.StartsWith("get", StringComparison.OrdinalIgnoreCase) Then
					attributeName = attribute.Name.LocalName
					callback = attribute.Value
					Return True
				Else
					attributeName = Nothing
					callback = Nothing
					Return False
				End If
			End Function

			Private Function InvokeMethod(name As String, ParamArray args() As Object) As Object
				Return extensibility.
					GetType().
					GetMethod(name, BindingFlags.Public Or BindingFlags.Instance).
					Invoke(extensibility, New Object() {Me}.Concat(args).ToArray())
			End Function
		End Class

	End Class

End Namespace