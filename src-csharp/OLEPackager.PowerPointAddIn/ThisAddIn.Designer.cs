#pragma warning disable 414

namespace OLEPackager.PowerPointAddIn
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;
        private global::System.Object missing = global::System.Type.Missing;
        internal Microsoft.Office.Interop.PowerPoint.Application Application;

        public ThisAddIn(global::Microsoft.Office.Tools.Factory factory, global::System.IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            Globals.Factory = factory;
        }

        protected override void Initialize()
        {
            base.Initialize();
            Application = GetHostItem<Microsoft.Office.Interop.PowerPoint.Application>(typeof(Microsoft.Office.Interop.PowerPoint.Application), "Application");
            Globals.ThisAddIn = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            InitializeCachedData();
            InitializeControls();
            InitializeComponents();
            InitializeData();
        }

        protected override void FinishInitialization()
        {
            InternalStartup();
            OnStartup();
        }

        protected override void InitializeDataBindings()
        {
            BeginInitialization();
            BindToData();
            EndInitialization();
        }

        private void InitializeCachedData()
        {
            if (DataHost != null && DataHost.IsCacheInitialized)
            {
                DataHost.FillCachedData(this);
            }
        }

        private void InitializeData() { }
        private void BindToData() { }
        private void StartCaching(string memberName) { DataHost.StartCaching(this, memberName); }
        private void StopCaching(string memberName) { DataHost.StopCaching(this, memberName); }
        private bool IsCached(string memberName) { return DataHost.IsCached(this, memberName); }

        private void BeginInitialization()
        {
            BeginInit();
            CustomTaskPanes.BeginInit();
        }

        private void EndInitialization()
        {
            CustomTaskPanes.EndInit();
            EndInit();
        }

        private void InitializeControls()
        {
            CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
        }

        private void InitializeComponents() { }
        private bool NeedsFill(string memberName) { return DataHost.NeedsFill(this, memberName); }

        protected override void OnShutdown()
        {
            CustomTaskPanes.Dispose();
            base.OnShutdown();
        }
    }

    internal sealed partial class Globals
    {
        private static ThisAddIn thisAddIn;
        private static global::Microsoft.Office.Tools.Factory factory;
        private static ThisRibbonCollection thisRibbonCollection;

        private Globals() { }

        internal static ThisAddIn ThisAddIn
        {
            get { return thisAddIn; }
            set
            {
                if (thisAddIn != null) { throw new global::System.NotSupportedException(); }
                thisAddIn = value;
            }
        }

        internal static global::Microsoft.Office.Tools.Factory Factory
        {
            get { return factory; }
            set
            {
                if (factory != null) { throw new global::System.NotSupportedException(); }
                factory = value;
            }
        }

        internal static ThisRibbonCollection Ribbons
        {
            get
            {
                if (thisRibbonCollection == null)
                {
                    thisRibbonCollection = new ThisRibbonCollection(factory.GetRibbonFactory());
                }

                return thisRibbonCollection;
            }
        }
    }

    internal sealed partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonCollectionBase
    {
        internal ThisRibbonCollection(global::Microsoft.Office.Tools.Ribbon.RibbonFactory factory) : base(factory) { }
    }
}
