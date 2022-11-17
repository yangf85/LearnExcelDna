using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms.Integration;
using System.Windows.Interop;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Learn
{
    public class Plugin : IExcelAddIn
    {
        public static Excel.Application App => ExcelDnaUtil.Application as Excel.Application;

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public void AutoOpen()
        {
            IntelliSenseServer.Install();
            App.WorkbookOpen += w =>
            {
                MenuManager.LoadMenu();
            };
            App.WorkbookBeforeClose += (Excel.Workbook w, ref bool flag) =>
            {
                MenuManager.UnloadMenu();
            };
        }
    }
}