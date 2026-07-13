using System;
using OLEPackager.Core;
using Office = Microsoft.Office.Core;

namespace OLEPackager.PowerPointAddIn
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
            PowerPointOlePackageService powerPointPackages = new PowerPointOlePackageService(Application);
            SelectedPowerPointObjectService selectedObjects = new SelectedPowerPointObjectService(
                Application,
                powerPointPackages,
                new PptxPackageService(new OlePackageExtractor()));
            service = new OlePackagerService(
                Application,
                new ZipPackageService(),
                powerPointPackages,
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
