using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Learn
{
    [ComVisible(true)]
    public partial class MainRibbon : ExcelRibbon
    {
        public IRibbonUI Ribbon { get; set; }

        public override string GetCustomUI(string RibbonID)
        {
            return ReadXmlRibbon();
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            Ribbon = ribbonUI;
        }

        private string ReadXmlRibbon()
        {
            var asm = Assembly.GetExecutingAssembly();
            var names = asm.GetManifestResourceNames();
            var current = names.FirstOrDefault(n => n.Contains("MainRibbon.xml"));
            if (current == null) { throw new Exception("MainRibbon.xml 不存在"); }
            using (var reader = new StreamReader(asm.GetManifestResourceStream(current)))
            {
                return reader.ReadToEnd();
            }
        }
    }
}