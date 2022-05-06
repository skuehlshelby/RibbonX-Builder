Imports RibbonFactory.Controls
Imports RibbonFactory.Utilities.Testing

<TestClass()> Public Class ValidationWithXsdTests

    <TestMethod()>
    Public Sub RibbonXWithShowLabelAttributeOnGroupElementIsInvalid()
        Dim invalidRibbonX As String = <?xml version="1.0" encoding="utf-8"?>

                                       <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
                                           <ribbon>
                                               <tabs>
                                                   <tab>
                                                       <group label="group1" showLabel="true">
                                                           <button id="button1" description="test"/>
                                                       </group>
                                                   </tab>
                                               </tabs>
                                           </ribbon>
                                       </customUI>.ToString()

        Dim errorLog As RibbonErrorLog = Ribbon.GetErrors(invalidRibbonX)

        Assert.IsFalse(errorLog.None)
    End Sub

End Class