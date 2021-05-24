Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums.MSO

Namespace Builders

    Public Class TabBuilder
        Inherits Builder(Of Tab)
        Implements ISetInsertionPoint(Of TabBuilder)
        Implements ISetKeyTip(Of TabBuilder)
        Implements ISetLabel(Of TabBuilder)
        Implements ISetVisibility(Of TabBuilder)

        Public Overrides Function Build(tag As Object) As Tab
            Return New Tab(attributes:=Attributes, tag:=tag)
        End Function

        Public Function Visible() As TabBuilder Implements ISetVisibility(Of TabBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As TabBuilder Implements ISetVisibility(Of TabBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As TabBuilder Implements ISetVisibility(Of TabBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As TabBuilder Implements ISetVisibility(Of TabBuilder).Invisible
            AddVisible(visible:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithLabel(label As String) As TabBuilder Implements ISetLabel(Of TabBuilder).WithLabel
            AddLabel(label:= label)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As TabBuilder Implements ISetKeyTip(Of TabBuilder).WithKeyTip
            AddKeyTip(keyTip:= keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As TabBuilder Implements ISetKeyTip(Of TabBuilder).WithKeyTip
            AddKeyTip(keyTip:= keyTip, callback:= callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As TabBuilder Implements ISetInsertionPoint(Of TabBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As TabBuilder Implements ISetInsertionPoint(Of TabBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As TabBuilder Implements ISetInsertionPoint(Of TabBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As TabBuilder Implements ISetInsertionPoint(Of TabBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class

End NameSpace