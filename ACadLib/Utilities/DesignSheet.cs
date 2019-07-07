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
        /// Dictionary of Named Ranges
        /// </summary>
        public Dictionary<string, Range> XlWorkbookNames { get; set; }

        /// <summary>
        /// Static class containing the Named Ranges in this Design Sheet
        /// </summary>
        public static class NamedRanges
        {
            /// <summary>
            /// The whole Sheet of PipeDataXlOut
            /// </summary>
            public const string PipeDataXlOut = "PipeDataXlOut";

            public const string PipeDataXlIn = "PipeDataXlIn";

            public const string PipeDataXlInEndInvert = "PipeDataXlIn.EndInvert";

            public const string PipeDataXlInFrom = "PipeDataXlIn.From";

            public const string PipeDataXlInTo = "PipeDataXlIn.To";

            public const string PipeDataXlInHandle = "PipeDataXlIn.Handle";

            public const string PipeDataXlInInnerDiameter = "PipeDataXlIn.InnerDiameter";

            public const string PipeDataXlOutSlope = "PipeDataXlOut.Slope";

            public const string PipeDataXlOutHandle = "PipeDataXlOut.Handle";

            public const string PipeDataXlOutStartInv = "PipeDataXlOut.StartInvert";

            public const string PipeDataXlOutEndInv = "PipeDataXlOut.EndInvert";
        }

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

            // Get and Set the Workbook names
            XlWorkbookNames = new Dictionary<string, Range>();
            foreach ( Name name in XlWorkbook.Names )
            {
                XlWorkbookNames.Add(name.Name, name.RefersToRange);
            }

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
