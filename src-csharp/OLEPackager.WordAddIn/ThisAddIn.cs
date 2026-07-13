using System;
using System.Reflection;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;

namespace OLEPackager.WordAddIn
{
    public partial class ThisAddIn
    {
        private OlePackagerService service;

        internal OlePackagerService Service
        {
            get { return service; }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OlePackagerRibbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ZipPackageService zipPackages = new ZipPackageService();
            WordOlePackageService wordPackages = new WordOlePackageService(Application);
            SelectedWordObjectService selectedObjects = new SelectedWordObjectService(
                Application,
                wordPackages,
                new FlatOpcPackageService(),
                new OlePackageExtractor());
            service = new OlePackagerService(
                Application,
                zipPackages,
                wordPackages,
                selectedObjects);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            service = null;
        }

        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
