using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ACadLib.Utilities;
using ACadLib.Views;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Microsoft.Office.Interop.Excel;
using DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject;
using XlApplication = Microsoft.Office.Interop.Excel.Application;

namespace ACadLib
{
    /// <summary>
    /// Static Model class for the AutoSheet Command
    /// </summary>
    public static class AutoSheet
    {
        #region Private Fields

        /// <summary>
        /// The Main window for the AutoSheet Command
        /// </summary>
        private static AutoSheetWindow _autoSheetMainWindow;

        #endregion

        #region Public Methods

        /// <summary>
        /// Get a list of Pipe Networks from the AutoCAD Database
        /// </summary>
        /// <returns>A List of Networks</returns>
        public static List<Network> GetPipeNetworks()
        {
            var pipeNetworkIds = BootstrapApp.CivilDoc.GetPipeNetworkIds();

            var pipeNetworks = new List<Network>(pipeNetworkIds.Count);

            using (var ts = BootstrapApp.TransManager.StartTransaction())
            {
                foreach (ObjectId networkId in pipeNetworkIds)
                    pipeNetworks.Add(ts.GetObject(networkId, OpenMode.ForRead) as Network);
            }

            return pipeNetworks;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// The Current Design Sheet
        /// </summary>
        public static DesignSheet DesignSheet { get; set; }

        #endregion

        #region AutoCAD Commands

        /// <summary>
        /// Start the AutoSheet Application and Open a new
        /// <see cref="AutoSheetWindow"/>
        /// </summary>
        public static void StartApplication()
        {
            CreateWindow();
            ACadLogger.Log("AutoSheet Window Opened");
        }

        /// <summary>
        /// Opens a design Sheet given a path
        /// </summary>
        /// <param name="filePath">The path to the design sheet</param>
        public static void OpenDesignSheet(string filePath)
        {
            // Check to make sure that the workbook isn't already open
            if ( DesignSheet != null )
            {
                // see if it is closed but the COM objects
                // were not cleaned up properly
                try
                {
                    var test = DesignSheet.XlApp.Worksheets;
                }
                catch ( System.Exception )
                {
                    // Dispose the Design Sheet and reopen
                    DesignSheet.Dispose();
                    DesignSheet = null;
                    OpenDesignSheet(filePath);
                }

                // Since the Workbook is already open return
                return;
            }
            try
            {
                DesignSheet = new DesignSheet(filePath, "PipeDataXlOut", "PipeDataXlIn");
            }
            catch (COMException e)
            {
                DesignSheet = null;
                ACadLogger.Log($"Design Sheet could not be opened: {e}");
            }
        }

        public static void ExportPipeData(Network pipeNetwork)
        {
            if (DesignSheet == null) return;

            var xlIn = DesignSheet.PipeDataXlIn;

            const char handleRow = 'A';
            const char fromRow = 'B';
            var colNumber = 2;

            var pipesIds = pipeNetwork.GetPipeIds();


            using (var ts = BootstrapApp.TransManager.StartTransaction())
            {
                foreach ( ObjectId id in pipesIds )
                {
                    var pipe = ts.GetObject(id, OpenMode.ForRead) as Pipe;

                    xlIn.Range[$"{handleRow}{colNumber}"].Value2 = pipe.Handle.Value;
                    // xlIn.Range[$"{fromRow}{colNumber}"].Value2 = pipe.StartStructureId;
                    colNumber++;
                }
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Set the value of the main window to null so that only
        /// one can be created
        /// </summary>
        /// <param name="sender">The sender of the event</param>
        /// <param name="e">Any Event arguments</param>
        private static void AutoSheetMainWindowOnClosed(object sender, EventArgs e)
        {
            _autoSheetMainWindow = null;

            if ( DesignSheet == null ) return;
            DesignSheet.Quit();
            DesignSheet = null;

        }

        /// <summary>
        /// Create a new <see cref="AutoSheetWindow"/> Window
        /// </summary>
        private static void CreateWindow()
        {
            if (_autoSheetMainWindow == null)
            {
                _autoSheetMainWindow = new AutoSheetWindow();
                _autoSheetMainWindow.Closed += AutoSheetMainWindowOnClosed;
            }
            else
                ACadLogger.Log("Window Already Exists");
        }

        #endregion
    }
}
