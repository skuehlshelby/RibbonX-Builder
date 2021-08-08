Imports System.Reflection
Imports Microsoft.Office.Core
Imports RibbonFactory

Namespace Stubs

    ''' <summary>
    ''' A test class for simulating requests from Excel.
    ''' </summary>
    Public Class Excel
        Implements IRibbonUI
        Private ReadOnly _extensibility As IRibbonExtensibility
        Private ReadOnly _ribbonX As XDocument

        Public Sub New(ribbonExtensibility As IRibbonExtensibility)
            _ribbonX = XDocument.Parse(ribbonExtensibility.GetCustomUI(String.Empty))
            _extensibility = ribbonExtensibility
        End Sub

        ''' <summary>
        ''' Simulates a click on a ribbon element, as performed by a user in Excel.
        ''' </summary>
        ''' <param name="element"></param>
        Public Sub Click(element As RibbonElement)
            Dim controlXML As XElement = GetElementByID(element.ID)
            Dim method As MethodInfo = GetMethodFromControl(controlXML, "onAction")

            method.Invoke(_extensibility, new Object(){GetRibbonControl(controlXML)}) 'Send a faked IRibbonControl to the onAction method.
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

        ''' <summary>
        ''' Creates a fake IRibbonControl object, for purposes of faking requests from Excel.
        ''' Real requests from Excel will be very similar.
        ''' </summary>
        ''' <param name="control">The control to turn into an IRibbonControl</param>
        ''' <returns></returns>
        Private Shared Function GetRibbonControl(control As XElement) As IRibbonControl
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

        ''' <summary>
        ''' Simulates an Excel request for updated control information.
        ''' </summary>
        ''' <param name="controlID">The id on the control </param>
        Public Sub InvalidateControl(controlID As String) Implements IRibbonUI.InvalidateControl
            Dim controlXML As XElement = GetElementByID(controlID)
            Dim fakedIRibbonControl As IRibbonControl = GetRibbonControl(controlXML) 'This will be passed to the IRibbonExtensibility object, which houses all the callbacks pointed to in the XML.

            For Each attribute As XAttribute In controlXML.Attributes() 'Iterate through the attributes on the control, looking for ones that point to methods.
                If IsMethod(attribute) Then
                    Dim method As MethodInfo = GetMethodFromControl(controlXML, attribute.Name.LocalName)

                    If method.GetParameters().Length = 1 Then
                        method.Invoke(_extensibility, New Object(){fakedIRibbonControl})
                    End If
                End If
            Next
        End Sub

        Private Shared Function IsMethod(attribute As XAttribute) As Boolean
            Return attribute.Name.LocalName.StartsWith("get")
        End Function

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