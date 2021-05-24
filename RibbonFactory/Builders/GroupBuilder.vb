Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace Builders
    Public Class GroupBuilder
        Inherits Builder(Of Group)
        Implements ISetInsertionPoint(Of GroupBuilder)
        Implements ISetVisibility(Of GroupBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder)
        Implements ISetKeyTip(Of GroupBuilder)
        Implements ISetImage(Of GroupBuilder)

        Public Overrides Function Build(tag As Object) As Group
            Return New Group(Attributes, tag)
        End Function

        Public Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As GroupBuilder Implements ISetImage(Of GroupBuilder).WithImage
            AddImage(image:= image, callback:= callback)
            Return Me
        End Function

        Public Function WithImage(image As Drawing.Bitmap, callback As FromControl(Of IPictureDisp)) As GroupBuilder Implements ISetImage(Of GroupBuilder).WithImage
            AddImage(image:= image, callback:= callback)
            Return Me
        End Function

        Public Function Visible() As GroupBuilder Implements ISetVisibility(Of GroupBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As GroupBuilder Implements ISetVisibility(Of GroupBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As GroupBuilder Implements ISetVisibility(Of GroupBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As GroupBuilder Implements ISetVisibility(Of GroupBuilder).Invisible
            AddVisible(visible:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As GroupBuilder Implements ISetKeyTip(Of GroupBuilder).WithKeyTip
            AddKeyTip(keyTip:= keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As GroupBuilder Implements ISetKeyTip(Of GroupBuilder).WithKeyTip
            AddKeyTip(keyTip:= keyTip, callback:= callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As GroupBuilder Implements ISetInsertionPoint(Of GroupBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As GroupBuilder Implements ISetInsertionPoint(Of GroupBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As GroupBuilder Implements ISetInsertionPoint(Of GroupBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As GroupBuilder Implements ISetInsertionPoint(Of GroupBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As GroupBuilder Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder).WithLabel
            AddLabel(label:= label)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As GroupBuilder Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder).WithLabel
            AddLabel(label:= label, callback:= callback)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label, callback:= callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As GroupBuilder Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As GroupBuilder Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip, callback:= callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As GroupBuilder Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As GroupBuilder Implements ISetLabelScreenTipAndSuperTip(Of GroupBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip, callback:= callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As GroupBuilder Implements ISetImage(Of GroupBuilder).WithImage
            AddImage(image:= image)
            Return Me
        End Function
    End Class
End NameSpace