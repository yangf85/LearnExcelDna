using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Learn
{
    public partial class MainRibbon : ExcelRibbon
    {
        public void OnFindButtonClicked(IRibbonControl control)
        {
            Excel.Range source = Plugin.App.InputBox(Prompt: "选择数据源", Type: 8);
            Excel.Range refer = Plugin.App.InputBox(Prompt: "选择参考", Type: 8);

            var random = new Random();
            foreach (Excel.Range cell1 in source)
            {
                string text1 = cell1.Text;
                foreach (Excel.Range cell2 in refer)
                {
                    string text2 = cell2.Text;
                    var index = random.Next(0, 255 * 255 * 255);
                    if (text1.Contains(text2))
                    {
                        cell1.Interior.Color = index;
                        cell2.Interior.Color = index;
                    }
                }
            }
        }

        public void OnShowPanePressed(IRibbonControl control, bool flag)
        {
            PaneManager.Visible(flag);
        }
    }
}