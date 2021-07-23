Imports System.Reflection
Imports Microsoft.Office.Core
Imports RibbonFactory

Namespace Stubs

    Public Class Excel
        Implements IRibbonUI
        Private ReadOnly _extensibility As IRibbonExtensibility
        Private ReadOnly _ribbonX As XDocument

        Public Sub New(ribbonExtensibility As IRibbonExtensibility)
            _ribbonX = XDocument.Parse(ribbonExtensibility.GetCustomUI(String.Empty))
            _extensibility = ribbonExtensibility
        End Sub

        Public Sub Click(element As RibbonElement)
            Dim control As XElement = GetElementByID(element.ID)
            Dim method As MethodInfo = GetMethodFromControl(control, "onAction")

            method.Invoke(_extensibility, new Object(){GetRibbonControl(control)})
        End Sub

        Private Function GetElementByID(id As String) As XElement
            Return _ribbonX.Root.Descendants() _
                .Single(Function(e) e.Attribute("id") IsNot Nothing AndAlso e.Attribute("id").Value = id)
        End Function

        Private Function GetMethodFromControl(control As XElement, methodName As String) As MethodInfo
            return _extensibility _
                .GetType() _
                .GetMethod(control.Attribute(methodName).Value, BindingFlags.Public Or BindingFlags.Instance)
        End Function

        Private Shared Function GetRibbonControl(control As XElement) As RibbonControl
            Dim id As String = control.Attribute(NameOf(id)).Value
            Dim tag As String

            If control.Attribute(NameOf(tag)) IsNot Nothing Then
                tag = control.Attribute(NameOf(tag)).Value
            Else
                tag = Nothing
            End If

            Return New RibbonControl(id, Nothing, tag)
        End Function

        Public Sub Invalidate() Implements IRibbonUI.Invalidate
            'Do Nothing
        End Sub

        Public Sub InvalidateControl(controlID As String) Implements IRibbonUI.InvalidateControl
            Dim control As XElement = GetElementByID(controlID)
            Dim stub As RibbonControl = GetRibbonControl(control)

            For Each attribute As XAttribute In control.Attributes()
                If attribute.Name.LocalName.StartsWith("get") Then
                    Dim method As MethodInfo = GetMethodFromControl(control, attribute.Name.LocalName)

                    If method.GetParameters().Length = 1 Then
                        method.Invoke(_extensibility, New Object(){stub})
                    End If
                End If
            Next
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
    End Class
End NameSpace