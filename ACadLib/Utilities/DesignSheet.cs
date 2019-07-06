using System;
using System.Collections.Generic;
using System.IO;
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
        /// <summary>
        /// The Excel Application
        /// </summary>
        public XlApplication XlApp;

        /// <summary>
        /// The Workbook
        /// </summary>
        public Workbook XlWorkbook;

        /// <summary>
        /// Worksheet that AutoCAD Will read from
        /// </summary>
        public Worksheet PipeDataXlOut;

        /// <summary>
        /// Worksheet that AutoCAD Will write to
        /// </summary>
        public Worksheet PipeDataXlIn;

        /// <summary>
        /// Default Constructor
        /// </summary>
        /// <param name="filename">Workbook filename</param>
        /// <param name="pipeDataXlOut">The pipe data out sheet name</param>
        /// <param name="pipeDataXlIn">The pipe data in sheet name</param>
        public DesignSheet(string filename, string pipeDataXlOut, string pipeDataXlIn)
        {
            if ( !File.Exists(filename) || Path.GetExtension(filename) != ".xlsx" )
            {
                throw new COMException();
            }

            XlApp = new XlApplication()
            {
                Visible = true
            };

            XlWorkbook = null;
            PipeDataXlOut = null;
            PipeDataXlIn = null;

            XlWorkbook = XlApp.Workbooks.Open(filename);

            // Get the worksheets
            PipeDataXlOut = XlWorkbook.Worksheets[pipeDataXlOut];
            PipeDataXlIn = XlWorkbook.Worksheets[pipeDataXlIn];

            if (XlWorkbook == null || PipeDataXlOut == null || PipeDataXlIn == null)
                throw new COMException();

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
            if ( PipeDataXlOut != null )
                Marshal.FinalReleaseComObject(PipeDataXlOut);

            if ( PipeDataXlIn == null )
                Marshal.FinalReleaseComObject(PipeDataXlIn);

            if (XlWorkbook != null)
                Marshal.FinalReleaseComObject(XlWorkbook);

            if (XlApp != null)
                Marshal.FinalReleaseComObject(XlApp);

        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }
    }
}
