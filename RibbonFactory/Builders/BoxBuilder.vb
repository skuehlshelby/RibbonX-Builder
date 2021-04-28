Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums

Namespace Builders
    
    Public Class BoxBuilder
        Inherits Builder(Of Box)
        Implements ISetOrientation(Of BoxBuilder)
        Implements ISetVisibility(Of BoxBuilder)
        
        Public Overrides Function Build() As Box
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As Box
            Return New Box(Attributes, tag)
        End Function

        Public Function Horizontal() As BoxBuilder Implements ISetOrientation(Of BoxBuilder).Horizontal
            AddOrientation(BoxStyle.vertical)
            Return Me
        End Function

        Public Function Vertical() As BoxBuilder Implements ISetOrientation(Of BoxBuilder).Vertical
            AddOrientation(BoxStyle.horizontal)
            Return Me
        End Function
        
        Public Function Visible() As BoxBuilder Implements ISetVisibility(Of BoxBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As BoxBuilder Implements ISetVisibility(Of BoxBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As BoxBuilder Implements ISetVisibility(Of BoxBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As BoxBuilder Implements ISetVisibility(Of BoxBuilder).Invisible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function
        
        Public Shared Function Horizontal(ParamArray items() As RibbonElement) As Box
            Dim horizontalBox As Box = new BoxBuilder().Horizontal().Build(tag:= Nothing)
            
            For Each item As RibbonElement In items
                horizontalBox.Add(item)
            Next item
            
            Return horizontalBox
        End Function
        
        Public Shared Function Vertical(ParamArray items() As RibbonElement) As Box
            Dim verticalBox As Box = new BoxBuilder().Vertical().Build(tag:= Nothing)
            
            For Each item As RibbonElement In items
                verticalBox.Add(item)
            Next item
            
            Return verticalBox
        End Function
        
    End Class
    
End NameSpace