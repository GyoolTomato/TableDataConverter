using System.Threading;

namespace TableDataConverter
{
    internal static class Program
    {
        static Mutex _mutex = null;

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            var isNew = false;
            _mutex = new Mutex(true, typeof(Program).FullName, out isNew);

            if (isNew == false)
            {
                MessageBox.Show("Process is already running...");
                return;
            }

            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}