using System;
using System.Diagnostics;
using System.IO;
using RedGate.SIPFrameworkShared;

namespace EdtDbProfiler.Ecosystem
{
    class EdtDbProfiler : ISsmsAddin4
    {
        public string Version => "2.0";
        public string Description => "Ecosystem integration for ExpressProfiler";
        public string Name => "ExpressProfiler";
        public string Author => "ExpressProfiler";
        public string Url => "https://expressprofiler.codeplex.com/";

        internal static ISsmsFunctionalityProvider6 m_Provider;
        public void OnLoad(ISsmsExtendedFunctionalityProvider provider)
        {
            m_Provider = (ISsmsFunctionalityProvider6)provider;
            m_Provider.AddToolbarItem(new ExecuteExpressProfiler());
            var command = new ExecuteExpressProfiler();
            m_Provider.AddToolsMenuItem(command);
        }


        public void OnNodeChanged(ObjectExplorerNodeDescriptorBase node)
        {
        }

        public void OnShutdown(){}
    }

    public class ExecuteExpressProfiler :  ISharedCommand
    {
        public void Execute()
        {
            string param ="";
            IConnectionInfo2 con;
            if(EdtDbProfiler.m_Provider.ObjectExplorerWatcher.TryGetSelectedConnection(out con))
            {
                string server = con.Server;
                string user = con.UserName;
                string password = con.Password;
                bool trusted = con.IsUsingIntegratedSecurity;
                param = trusted ? String.Format("-server \"{0}\"",server) : String.Format("-server \"{0}\" -user \"{1}\" -password \"{2}\"", server,user,password);
            }
            string root = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string profiler = Path.Combine(root, "ExpressProfiler\\ExpressProfiler.exe");
            Process.Start(profiler, param);
        }

        private readonly ICommandImage m_CommandImage = new CommandImageForEmbeddedResources(typeof(ExecuteExpressProfiler).Assembly, "ExpressProfiler.Ecosystem.Resources.Icon.png");
        public string Name => "ExpressProfilerExecute";
        public string Caption => "ExpressProfiler";
        public string Tooltip => "Execute ExpressProfiler";
        public ICommandImage Icon => m_CommandImage;
        public string[] DefaultBindings => new string[] { };
        public bool Visible => true;
        public bool Enabled => true;
    }
}
