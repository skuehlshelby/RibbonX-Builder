Imports System.Drawing
Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Controls
Imports RibbonX.Images
Imports RibbonX.InternalApi

Namespace Builders
    Friend NotInheritable Class ItemBuilder
        Inherits BuilderBase(Of Item)
        Implements IItemBuilder

        Public Sub New()
            MyBase.New()
            Dim id As IRibbonProperty = ControlProperties.Single(Function(p) p.Category.Equals(Category.IdType))
            ControlProperties.Clear()
            ControlProperties.Add(
                    id,
                    RibbonProperty.Create(Name.Label, Category.Label, String.Empty, NameOf(FakeCallback)),
                    RibbonProperty.Create(Name.Screentip, Category.ScreenTip, String.Empty, NameOf(FakeCallback)),
                    RibbonProperty.Create(Name.Supertip, Category.SuperTip, String.Empty, NameOf(FakeCallback)),
                    RibbonProperty.Create(Name.Image, Category.Image, RibbonImage.Create(New Bitmap(1, 1)), NameOf(FakeCallbackII)))
        End Sub

        Public Overrides Function Build() As Item
            Return New Item(ControlProperties, Tag)
        End Function

        Public Function WithTag(tag As Object) As IItemBuilder Implements ITag(Of IItemBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Private Function WithId(id As String) As IItemBuilder Implements IItemBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreentip As Boolean = True) As IItemBuilder Implements IItemBuilder.WithLabel
            LabelBase(label, AddressOf FakeCallback)
            If copyToScreentip Then
                ScreentipBase(label, AddressOf FakeCallback)
                ControlProperties.Forward(Category.ScreenTip, Category.Label)
            End If
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

        Public Function FromTemplate(template As IRibbonElement) As IItemBuilder Implements ITemplatable(Of IItemBuilder).FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IItemBuilder) = Nothing) As IItem
            Dim instance As ItemBuilder = New ItemBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

        Private Shared Function FakeCallback(_1 As IRibbonControl) As String
            Return String.Empty
        End Function

        Private Shared Function FakeCallbackII(_1 As IRibbonControl) As IPictureDisp
            Return Nothing
        End Function

    End Class

End Namespace