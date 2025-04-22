using System;
using System.Windows.Forms;

namespace NCreoConvertTool
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            //string sTempWorkFold = "";
            //for (int i = 0; i < args.Length; i++)
            //{
            //    if (i == 0)
            //        sTempWorkFold = args[i];
            //}

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
