Imports System.Drawing
Imports Microsoft.Office.Core
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Controls
Imports stdole

Namespace Builders
    Friend NotInheritable Class ItemBuilder
        Inherits BuilderBase(Of Item)
        Implements IItemBuilder

        Public Sub New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function WithLabel(label As String, Optional copyToScreentip As Boolean = True) As IItemBuilder Implements IItemBuilder.WithLabel
            LabelBase(label, AddressOf FakeCallback)

            If copyToScreentip Then ScreentipBase(label, AddressOf FakeCallback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IItemBuilder Implements IItemBuilder.WithScreenTip
            ScreentipBase(screenTip, AddressOf FakeCallback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IItemBuilder Implements IItemBuilder.WithSuperTip
            SupertipBase(superTip, AddressOf FakeCallback)
            Return Me
        End Function

        Private Function WithImage(image As Bitmap) As IItemBuilder Implements IItemBuilder.WithImage
            ImageBase(image, AddressOf FakeCallbackII)
            Return Me
        End Function

        Private Function WithImage(image As Icon) As IItemBuilder Implements IItemBuilder.WithImage
            ImageBase(image, AddressOf FakeCallbackII)
            Return Me
        End Function

        Private Function FakeCallback(control As IRibbonControl) As String
            Return String.Empty
        End Function

        Private Function FakeCallbackII(control As IRibbonControl) As IPictureDisp
            Return Nothing
        End Function

    End Class

End Namespace