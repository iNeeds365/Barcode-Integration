using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Globalization;
using System.Windows.Interop;

namespace OrderPack
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application, IDisposable
    {
        private Mutex m_Mutex;

        static public bool is_closing = false;
        private void Dispose(Boolean disposing)
        {
            if (disposing && (m_Mutex != null))
            {
                m_Mutex.ReleaseMutex();
                m_Mutex.Close();
                m_Mutex = null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected override void OnStartup(StartupEventArgs e)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Boolean mutexCreated;
            String mutexName = String.Format(CultureInfo.InvariantCulture, "Local\\{{{0}}}{{{1}}}", assembly.GetType().GUID, assembly.GetName().Name);

            m_Mutex = new Mutex(true, mutexName, out mutexCreated);

            if (!mutexCreated)
            {
                m_Mutex = null;
                is_closing = true;
                MessageBox.Show("Another instance is running.");
                Current.Shutdown();
                return;
            }

            base.OnStartup(e);

            //MainWindow window = new MainWindow();
            //MainWindow = window;
            //window.Show();

            //HwndSource.FromHwnd((new WindowInteropHelper(window)).Handle).AddHook(new HwndSourceHook(HandleMessages));
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Dispose();
            base.OnExit(e);
        }
    }
}
