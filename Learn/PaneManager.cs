using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn
{
    public class PaneManager
    {
        private static CustomTaskPane _pane;

        public static void Visible(bool isVisible)
        {
            _pane = CustomTaskPaneFactory.CreateCustomTaskPane(new PaneContent(), "自定义窗格");
            _pane.Visible = isVisible;
            _pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _pane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
        }
    }
}