using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using System.Reflection;
using System.Windows.Threading;
using System.Windows.Navigation;
using System.Diagnostics;
using System.Runtime.InteropServices;
using KB.Processes;
using KB.Configuration;
using KB.Utility;

namespace HTTPRequestScheduler
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            ProcessHelper.HideMainWindow();
            Ini.Default.CreateDefault(HTTPRequestScheduler.Properties.Resources.DefaultIni);
            Ini.Default.IniFile.WriteValues("" ,ProcessHelper.ArgsValues);
        }

        protected override async void OnStartup(StartupEventArgs e)
        {
            // Check Arguments Options
            if (ProcessHelper.IsArgumentExists("/help"))
            {
                Console.WriteLine(HTTPRequestScheduler.Properties.Resources.arguments);
                this.Shutdown();
            }

            Worker worker = new Worker();
            Ini.Default.LoadProperties(worker, string.Empty);

            MainWindow mw = (MainWindow)(this.MainWindow = new MainWindow(worker));

            if (!ProcessHelper.IsArgumentExists("/hide", "/h"))
                mw.Show();
            if (ProcessHelper.IsArgumentExists("/send", "/s"))
                await worker.Send();
            if (ProcessHelper.IsArgumentExists("/close", "/c"))
                this.Shutdown();
            else worker.StartScheduling();

            base.OnStartup(e);
        }

        void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs args)
        {
            Console.WriteLine();
            ConsoleHelper.WriteErrorLine("Crashed:");
            //args.Handled = true; // Prevent default unhandled exception processing
        }
    }
}
