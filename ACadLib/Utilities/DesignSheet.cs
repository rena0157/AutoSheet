using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using XlApplication = Microsoft.Office.Interop.Excel.Application;

namespace ACadLib.Utilities
{
    public class DesignSheet : IDisposable
    {
        public XlApplication XlApp;
        public Workbook XlWorkbook;
        public Worksheet PipeDataSheet;


        public DesignSheet(string filename, string worksheetName)
        {
            XlApp = new XlApplication()
            {
                Visible = true
            };

            XlWorkbook = null;
            PipeDataSheet = null;

            XlWorkbook = XlApp.Workbooks.Open(filename);
            PipeDataSheet = XlWorkbook.Worksheets[worksheetName];

            ACadLogger.Log("Pipe Data Sheet Opened");
        }

        ~DesignSheet()
        {
            ReleaseUnmanagedResources();
        }

        public void Quit()
        {
            XlWorkbook.Close();
            XlApp.Quit();
            Dispose();
        }

        private void ReleaseUnmanagedResources()
        {
            Marshal.FinalReleaseComObject(PipeDataSheet);
            Marshal.FinalReleaseComObject(XlWorkbook);
            Marshal.FinalReleaseComObject(XlApp);
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }
    }
}
