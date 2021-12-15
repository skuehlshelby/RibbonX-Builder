Imports System.Drawing
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities
Imports stdole

Namespace Builders
    Public NotInheritable Class ItemBuilder
        Private ReadOnly _id As String
        Private _label As String
        Private _screenTip As String
        Private _superTip As String
        Private _image As IPictureDisp
        Private _synchronizeLabelAndScreenTip As Boolean

        Public Sub New()
            _id = IdManager.GetID(Of Item)
            _label = String.Empty
            _screenTip = String.Empty
            _superTip = String.Empty
            _image = Nothing
            _synchronizeLabelAndScreenTip = False
        End Sub

        Public Function Build(Optional tag As Object = Nothing) As Item
            Return New Item(_id, _label, _screenTip, _superTip, _image, _synchronizeLabelAndScreenTip, tag)
        End Function

        Public Function WithLabel(label As String, Optional synchronizeLabelAndScreenTip As Boolean = True) As ItemBuilder
            _label = label

            _synchronizeLabelAndScreenTip = synchronizeLabelAndScreenTip

            If synchronizeLabelAndScreenTip Then
                _screenTip = _label
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ItemBuilder
            _screenTip = screenTip

            If _synchronizeLabelAndScreenTip Then
                _label = _screenTip
            End If

            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ItemBuilder
            _superTip = superTip
            Return Me
        End Function

        Public Function WithImage(image As Bitmap) As ItemBuilder
            _image = New AdapterForIPicture(image)
            Return Me
        End Function

        Public Function WithImage(image As Icon) As ItemBuilder
            _image = New AdapterForIPicture(image)
            Return Me
        End Function
    End Class
End NameSpace