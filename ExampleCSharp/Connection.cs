using RibbonX;
using RibbonX.Controls;
using RibbonX.Controls.BuiltIn;
using RibbonX.Images;
using RibbonX.Images.BuiltIn;
using System;
using System.Runtime.InteropServices;
using MessageBox = System.Windows.Forms.MessageBox;
using MessageBoxButtons = System.Windows.Forms.MessageBoxButtons;

namespace ExampleCSharp
{
    [ComVisible(true)]
    [Guid("68D5D1E1-A37E-4E5A-A706-D2EBF4FA927E")]
    [ProgId("ExampleCSharp.Connection")]
    public class Connection : StockRibbonBase
    {
        internal Button Button { get; }
        internal Group Group { get; }
        internal Tab Tab { get; }
        
        public Connection()
        {
            Button = new Button(b => b.Large()
                                      .WithLabel("Click Me!", GetLabel)
                                      .WithSuperTip("I Really Want To Be Clicked!", GetSuperTip)
                                      .WithImage(Common.SadFace, GetBuiltInImage)
                                      .RouteClickTo(OnAction)
                                      .OnClick((_) => MessageBox.Show(caption: "RibbonX C-Sharp Example", text: "Thank you!", buttons: MessageBoxButtons.OK),
                                      (btn) =>
                                      {
                                          using (btn.RefreshSuspension())
                                          {
                                              btn.Label = "Thank You";
                                              btn.SuperTip = "I really needed that.";
                                              btn.Image = RibbonImage.Create(Common.HappyFace);
                                          }
                                      }));
                                          
            Group = new Group(b => b.WithLabel("A Button In Need")
                                    .WithSuperTip("This group contains a button that needs your help."), 
                                    controls: new []{ Button });

            Tab = new Tab(b => b.WithLabel("C# Example")
                                .InsertAfter(Excel.TabHome), 
                                groups: new[] { Group });
        }

        protected override Ribbon BuildRibbon()
        {
            return new Ribbon(b => b.OnLoad(OnLoad), Tab);
        }
    }
}
