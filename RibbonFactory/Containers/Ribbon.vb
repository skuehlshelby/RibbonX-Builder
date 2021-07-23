Imports Microsoft.Office.Core
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes

Namespace Containers

    Public NotInheritable Class Ribbon

        Private Const BoilerPlate As String =
            "<?xml version=""1.0"" encoding=""utf-8"" ?>" _
            & vbNewLine & vbNewLine &
            "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" _
            & vbNewLine & vbTab & "<ribbon {0}>" _
            & vbNewLine & vbTab & vbTab & "<tabs>" _
            & vbNewLine & vbTab & vbTab & vbTab & "{1}" _
            & vbNewLine & vbTab & vbTab & "</tabs>" _
            & vbNewLine & vbTab & "</ribbon>" _
            & vbNewLine & "</customUI>"

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _tabs As List(Of Tab)
        Private ReadOnly _startFromScratch As Boolean
        Private ReadOnly _allElements As List(Of RibbonElement) = New List(Of RibbonElement)
        Private _ribbon As IRibbonUI

        Public Sub New(Optional startFromScratch As Boolean = False)
            _startFromScratch = startFromScratch
            _tabs = New List(Of Tab)
        End Sub

        Public Function GetElement(id As String) As RibbonElement
            For Each t As Tab In _tabs

                If t.ID = id Then
                    Return t
                End If

                For Each g As Group In t.Groups

                    If g.ID = id Then
                        Return g
                    End If

                    For Each c As RibbonElement In g.Controls
                        If c.ID = id Then
                            Return c
                        End If
                    Next c
                Next g
            Next t

            Throw New InvalidOperationException($"Item with ID '{id}' was not found.")
        End Function

        Public Function GetElement(Of T)(id As String) As T
            Return CType(CType(GetElement(id), Object), T)
        End Function

        Public Function GetChildItem(parentId As String, index As Integer) As DropdownItem
            Return GetElement(Of IEnumerable(Of DropdownItem))(parentId).ElementAt(index)
        End Function

        Public ReadOnly Property Tabs As IReadOnlyList(Of Tab)
            Get
                Return _tabs
            End Get
        End Property

        Public Function AddTab(tab As Tab) As Tab
            _tabs.Add(tab)
            AddHandler tab.ValueChanged, AddressOf HandleValueChanged
            Return tab
        End Function

        Private Sub HandleValueChanged(sender As Object, e As ValueChangedEventArgs)
            _ribbon.InvalidateControl(e.ID)
        End Sub

        Public Function Build() As String
            Return String.Format(BoilerPlate, $"startFromScratch=""{_startFromScratch.ToString().ToLower()}""", String.Join(vbNewLine, Tabs.Select(Function(T) T.XML)))
        End Function

        Public Sub AssignRibbonUI(ribbonUi As IRibbonUI)
            _ribbon = If(_ribbon, ribbonUi)
        End Sub

    End Class
End NameSpace