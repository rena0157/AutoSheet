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
        public static PipeDataSheet DataSheet { get; private set; }

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
            if ( DataSheet != null )
            {
                if ( DataSheet.IsReady() ) return;
                // Dispose the Design Sheet and reopen
                DataSheet.Dispose();
                DataSheet = null;
            }
            try
            {
                DataSheet 
                    = new PipeDataSheet(filePath, PipeDataSheet.PipeDataSheetName);
                if ( !DataSheet.IsReady() )
                    DataSheet = null;

            }
            catch (COMException)
            {
                DataSheet = null;
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
            if (DataSheet == null) return;

            // The current row/pipe number (Starts at 2 because of headers)
            var rowNumber = 2;

            // All of the pipe Ids from the network
            var pipesIds = pipeNetwork.GetPipeIds();

            // Access all of the data and place it into the excel sheet
            using (var ts = BootstrapApp.TransManager.StartTransaction())
            {
                foreach ( ObjectId id in pipesIds )
                {
                    var pipe = ts.GetObject(id, OpenMode.ForRead) as Pipe;

                    // If the Id is invalid then continue
                    if (pipe == null) continue;

                    // Get all ranges and cells
                    Range handleCell = DataSheet.HandleRange[rowNumber];
                    Range fromCell = DataSheet.FromRange[rowNumber];
                    Range toCell = DataSheet.ToRange[rowNumber];
                    Range lengthCell = DataSheet.LengthRange[rowNumber];
                    Range diameterCell = DataSheet.InnerDiameterRange[rowNumber];

                    handleCell.Value2 = pipe.Handle;

                    // Get start and End Structures
                    var startStructure = ts.GetObject(pipe.StartStructureId, OpenMode.ForRead) as Structure;
                    var endStructure = ts.GetObject(pipe.EndStructureId, OpenMode.ForRead) as Structure;

                    // Set Start and End Structure Names
                    fromCell.Value2 = startStructure == null ? "null" : startStructure.Name;
                    toCell.Value2 = endStructure == null ? "null" : endStructure.Name;

                    // Set the length and diameter values
                    lengthCell.Value2 = pipe.Length2DCenterToCenter;
                    diameterCell.Value2 = pipe.InnerDiameterOrWidth;

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
            if (DataSheet == null || !DataSheet.IsReady()) return;


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
                        // Get that start Invert from the DataSheet
                        double startInv = DataSheet
                                            .XlApp
                                            .WorksheetFunction
                                            .VLookup(handle, DataSheet.PipeDataRange, DataSheet.StartInvRange.Column, false) ?? 0;

                        // Get the End Invert from the DataSheet
                        double endInv = DataSheet
                                            .XlApp
                                            .WorksheetFunction
                                            .VLookup(handle, DataSheet.PipeDataRange, DataSheet.EndInvRange.Column, false) ?? 0;

                        // Set new Start and End Points for the Pipe
                        pipe.StartPoint = new Point3d(pipe.StartPoint.X, pipe.StartPoint.Y, startInv);
                        pipe.EndPoint = new Point3d(pipe.EndPoint.X,pipe.EndPoint.Y,endInv);

                        // Disconnect and Reconnect the Start Structure
                        var startStructureId = pipe.StartStructureId;
                        pipe.Disconnect(ConnectorPositionType.Start);
                        pipe.ConnectToStructure(ConnectorPositionType.Start, startStructureId, true);

                        // Disconnect and Reconnect the End Structure
                        var endStructureId = pipe.EndStructureId;
                        pipe.Disconnect(ConnectorPositionType.End);
                        pipe.ConnectToStructure(ConnectorPositionType.End, endStructureId, true);
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

            if ( DataSheet == null ) return;
            DataSheet.Dispose();
            DataSheet = null;

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
