using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OLEPackager.WordAddIn
{
    [ComVisible(true)]
    public sealed class OlePackagerRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonId)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OLEPackager.WordAddIn.Ribbon.OlePackagerRibbon.xml"))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void OnLoad(Office.IRibbonUI ribbonUi)
        {
            ribbon = ribbonUi;
        }

        public Bitmap InsertOLEPackageIcon(Office.IRibbonControl control)
        {
            return LoadBitmap("insert-ole-package-icon.png");
        }

        public Bitmap InsertImageFilesIcon(Office.IRibbonControl control)
        {
            return LoadBitmap("insert-image-files-icon.png");
        }

        public Bitmap ManageImageOleFilesIcon(Office.IRibbonControl control)
        {
            return LoadBitmap("manage-image-ole-files-icon.png");
        }

        public Bitmap UsageHelpIcon(Office.IRibbonControl control)
        {
            return LoadBitmap("usage-help-icon.png");
        }

        public void OnInsertPlotFolder(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Service.InsertPlotFolderAsOle();
        }

        public void OnInsertImageSupportFiles(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Service.InsertImageSupportFilesAsOle();
        }

        public void OnAttachSupportFilesToSelectedImage(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Service.AttachSupportFilesToSelectedImageAsOle();
        }

        public void OnShowUsage(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Service.ShowUsageHelp();
        }

        private static Bitmap LoadBitmap(string fileName)
        {
            string resourceName = "OLEPackager.WordAddIn.Assets." + fileName;
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidDataException("Missing Ribbon image resource: " + resourceName);
                }

                using (Bitmap source = new Bitmap(stream))
                {
                    return new Bitmap(source);
                }
            }
        }
    }
}
