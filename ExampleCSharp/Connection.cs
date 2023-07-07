using RibbonX;
using RibbonX.Controls.BuiltIn;
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
        internal IButton Button { get; private set; }
        internal IGroup Group { get; private set; }
        internal ITab Tab { get; private set; }

        protected override IRibbon BuildRibbon()
        {
            Button = RxApi.Button(b => b.Large()
              .WithLabel("Click Me!", GetLabel)
              .WithSuperTip("I Really Want To Be Clicked!", GetSuperTip)
              .WithImage(Common.SadFace, GetBuiltInImage)
              .OnClick(OnAction, click => click.Do(btn =>
              {
                  
                  using (btn.SuspendRefreshing())
                  {
                      btn.Label = "Thank You";
                      btn.SuperTip = "I really needed that.";
                      btn.Image = Common.HappyFace;
                  }

                  MessageBox.Show(caption: "RibbonX C-Sharp Example", text: "Thank you!", buttons: MessageBoxButtons.OK);
              })));

            Group = RxApi.Group(b => b.WithLabel("A Button In Need")
                                    .WithSuperTip("This group contains a button that needs your help.")
                                    .WithControls(Button));

            Tab = RxApi.Tab(b => b.WithLabel("C# Example")
                                .InsertAfter(Excel.TabHome)
                                .WithGroups(Group));

            return RxApi.Ribbon(b => b.OnLoad(OnLoad).WithTabs(Tab));
        }
    }
}
