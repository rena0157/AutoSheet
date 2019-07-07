using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using XlApplication = Microsoft.Office.Interop.Excel.Application;

namespace ACadLib.Utilities
{
    /// <summary>
    /// Represents a Design sheet in Excel
    /// </summary>
    public class DesignSheet : IDisposable
    {
        /// <summary>
        /// The Excel Application
        /// </summary>
        public readonly XlApplication XlApp;

        /// <summary>
        /// The Workbook
        /// </summary>
        private readonly Workbook _xlWorkbook;

        /// <summary>
        /// Worksheet that AutoCAD Will read from
        /// </summary>
        private readonly Worksheet _pipeDataXlOut;

        /// <summary>
        /// Worksheet that AutoCAD Will write to
        /// </summary>
        public readonly Worksheet PipeDataXlIn;

        /// <summary>
        /// Dictionary of Named Ranges
        /// </summary>
        public Dictionary<string, Range> XlWorkbookNames { get; }

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
            if ( filename == null )
            {
                MessageBox.Show("You must provide a filename", "AUTOSHEET ERROR", MessageBoxButton.OK);
                return;
            }

            if ( !File.Exists(filename) || Path.GetExtension(filename) != ".xlsx" )
            {
                MessageBox.Show($"Error Opening File: {filename}", "AUTOSHEET ERROR", MessageBoxButton.OK);
                return;
            }

            XlApp = new XlApplication()
            {
                Visible = true
            };

            _xlWorkbook = null;
            _pipeDataXlOut = null;
            PipeDataXlIn = null;

            _xlWorkbook = XlApp.Workbooks.Open(filename);

            // Get the worksheets
            _pipeDataXlOut = _xlWorkbook.Worksheets[pipeDataXlOut];
            PipeDataXlIn = _xlWorkbook.Worksheets[pipeDataXlIn];

            if (_xlWorkbook == null || _pipeDataXlOut == null || PipeDataXlIn == null)
                throw new COMException();

            // Get and Set the Workbook names
            XlWorkbookNames = new Dictionary<string, Range>();
            foreach ( Name name in _xlWorkbook.Names )
            {
                XlWorkbookNames.Add(name.Name, name.RefersToRange);
            }

            ACadLogger.Log("Pipe Data Sheet Opened");
        }

        /// <summary>
        /// Finalizer for this class
        /// </summary>
        ~DesignSheet()
        {
            ReleaseUnmanagedResources();
        }

        /// <summary>
        /// Returns true if the design sheet is ready
        /// and working properly
        /// </summary>
        /// <returns>Returns true if the design sheet is working properly</returns>
        public bool IsReady()
        {
            try
            {
                // Try to call the worksheets object
                var test = XlApp.Worksheets;
            }
            catch (System.Exception)
            {
                // If anything goes wrong then the sheet is
                // not working properly
                return false;
            }

            return true;
        }

        /// <summary>
        /// Dispose of this object releasing all of its
        /// unmanaged resources and COM objects
        /// </summary>
        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        /// <summary>
        /// Get the Excel Process of this object
        /// </summary>
        private Process ExcelProcess
        {
            get
            {
                GetWindowThreadProcessId(XlApp.Hwnd, out var id);
                return Process.GetProcessById(id);
            }
        }

        /// <summary>
        /// Release all of the unmanaged resources and COM objects
        /// </summary>
        private void ReleaseUnmanagedResources()
        {
            if ( _pipeDataXlOut != null )
                Marshal.FinalReleaseComObject(_pipeDataXlOut);

            if ( PipeDataXlIn != null )
                Marshal.FinalReleaseComObject(PipeDataXlIn);

            if (_xlWorkbook != null)
                Marshal.FinalReleaseComObject(_xlWorkbook);

            if ( XlApp == null ) return;

            ExcelProcess?.Kill();
            Marshal.FinalReleaseComObject(XlApp);
        }


    }
}
