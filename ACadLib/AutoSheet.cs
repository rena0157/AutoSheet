using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ACadLib.Utilities;
using ACadLib.Views;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
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
        public static IEnumerable<Network> GetPipeNetworks()
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
        public static DesignSheet DesignSheet { get; private set; }

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
                if ( DesignSheet.IsReady() ) return;
                // Dispose the Design Sheet and reopen
                DesignSheet.Dispose();
                DesignSheet = null;
            }
            try
            {
                DesignSheet 
                    = new DesignSheet(filePath, "PipeDataXlOut", "PipeDataXlIn");
                if ( !DesignSheet.IsReady() )
                    DesignSheet = null;

            }
            catch (COMException e)
            {
                DesignSheet = null;
            }
        }

        /// <summary>
        /// Exports data to the Design Sheet in the XlIn Sheet
        /// </summary>
        /// <param name="pipeNetwork">The pipe network that is to be
        /// exported</param>
        public static void ExportPipeData(Network pipeNetwork)
        {
            // If there is no active design sheet then nothing can be
            // exported
            if (DesignSheet == null) return;

            // Get the right tab from the design sheet
            var xlIn = DesignSheet.PipeDataXlIn;

            // Column Constants
            const char handleColumn = 'A';
            const char fromColumn = 'B';
            const char toColumn = 'C';
            const char lengthColumn = 'D';
            const char slopeColumn = 'E';
            const char innerDiaColumn = 'F';
            const char startInvColumn = 'G';
            const char endInvColumn = 'H';

            // The current row/pipe number
            var rowNumber = 2;

            // All of the pipe Ids from the network
            var pipesIds = pipeNetwork.GetPipeIds();

            // Access all of the data and place it into the excel sheet
            using (var ts = BootstrapApp.TransManager.StartTransaction())
            {
                foreach ( ObjectId id in pipesIds )
                {
                    var pipe = ts.GetObject(id, OpenMode.ForRead) as Pipe;
                    if (pipe != null)
                        xlIn.Range[$"{handleColumn}{rowNumber}"].Value2 = pipe.Handle.Value;
                    
                    var startStructure = ts.GetObject(pipe.StartStructureId, OpenMode.ForRead) as Structure;
                    if (startStructure != null)
                        xlIn.Range[$"{fromColumn}{rowNumber}"].Value2 = startStructure.Name;

                    var endStructure = ts.GetObject(pipe.EndStructureId, OpenMode.ForRead) as Structure;
                    if (endStructure != null)
                        xlIn.Range[$"{toColumn}{rowNumber}"].Value2 = endStructure.Name;

                    xlIn.Range[$"{lengthColumn}{rowNumber}"].Value2 = pipe.Length2DCenterToCenter;
                    xlIn.Range[$"{slopeColumn}{rowNumber}"].Value2 = pipe.Slope;
                    xlIn.Range[$"{innerDiaColumn}{rowNumber}"].Value2 = pipe.InnerDiameterOrWidth;
                    xlIn.Range[$"{startInvColumn}{rowNumber}"].Value2 = pipe.StartPoint.Z;
                    xlIn.Range[$"{endInvColumn}{rowNumber}"].Value2 = pipe.EndPoint.Z;

                    // Increment to the next row for the next pipe
                    rowNumber++;
                }
            }
        }

        /// <summary>
        /// Imports Pipe Data into Excel
        /// </summary>
        /// <param name="network">The network that will be imported</param>
        public static void ImportPipeData(Network network)
        {
            // If the design sheet is null or not ready return
            if (DesignSheet == null || !DesignSheet.IsReady()) return;

            // Ranges
            var sheetRange = DesignSheet.XlWorkbookNames[DesignSheet.NamedRanges.PipeDataXlOut];
            var startInvCol = DesignSheet.XlWorkbookNames[DesignSheet.NamedRanges.PipeDataXlOutStartInv];
            var endInvCol = DesignSheet.XlWorkbookNames[DesignSheet.NamedRanges.PipeDataXlOutEndInv];

            using ( var ts = BootstrapApp.TransManager.StartTransaction() )
            {
                BootstrapApp.ActiveDocument.LockDocument();
                var pipeIds = network.GetPipeIds();

                foreach ( ObjectId pipeId in pipeIds )
                {
                    var pipe = ts.GetObject(pipeId, OpenMode.ForWrite) as Pipe;

                    if (pipe == null) continue;

                    var handle = pipe.Handle.Value;

                    try
                    {
                        double startInv =
                            DesignSheet.XlApp.WorksheetFunction.VLookup(handle, sheetRange, startInvCol.Column, false) ?? 0;

                        double endInv =
                            DesignSheet.XlApp.WorksheetFunction.VLookup(handle, sheetRange, endInvCol.Column, false) ??
                            0;

                        pipe.StartPoint = new Point3d(pipe.StartPoint.X, pipe.StartPoint.Y, startInv);
                        pipe.EndPoint = new Point3d(pipe.EndPoint.X,pipe.EndPoint.Y,endInv);

                        // TODO: Need to reconnect the structures to the pipes
                    }
                    catch ( COMException )
                    {
                    }
                }

                // Commit the changes to the database
                ts.Commit();
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
            DesignSheet.Dispose();
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
