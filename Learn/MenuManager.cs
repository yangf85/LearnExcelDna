using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Learn
{
    public class MenuManager
    {
        private static Office.CommandBarButton _button1;
        private static Office.CommandBarButton _button2;
        private static Office.CommandBarPopup _control;

        public static void LoadMenu()
        {
            var bar = Plugin.App.CommandBars["Cell"];
            _control = bar.Controls.Add(Type: Office.MsoControlType.msoControlPopup, Temporary: true) as Office.CommandBarPopup;
            _control.Caption = "扩展菜单";
            _button1 = _control.Controls.Add(Type: Office.MsoControlType.msoControlButton, Temporary: true) as Office.CommandBarButton;
            _button1.Caption = "按钮1";
            _button1.Tag = "button1";
            _button1.FaceId = 0417;
            _button1.Click += Button_Click;
            _button2 = _control.Controls.Add(Type: Office.MsoControlType.msoControlButton, Temporary: true) as Office.CommandBarButton;
            _button2.Caption = "按钮2";
            _button2.FaceId = 0609;
            _button2.Click += Button_Click;
            _button2.Tag = "button2";
        }

        public static void UnloadMenu()
        {
            var bar = Plugin.App.CommandBars["Cell"];
            foreach (Office.CommandBarControl control in bar.Controls)
            {
                if (control.Caption == "扩展菜单")
                {
                    control.Delete();
                }
            }
        }

        private static void Button_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            switch (Ctrl.Tag)
            {
                case "button1":
                    Plugin.App.Dialogs[Excel.XlBuiltInDialog.xlDialogColorPalette].Show();
                    break;

                case "button2":
                    Plugin.App.Selection.Clear();
                    break;

                default:
                    break;
            }
        }
    }
}