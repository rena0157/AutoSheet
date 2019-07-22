// AutoSheet.cs
// By: Adam Renaud
// Created: 2019-07-21

using ACadLib.Exceptions;
using ACadLib.Utilities;
using ACadLib.Views;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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
                pipeNetworks.AddRange
                (
                    from ObjectId networkId in pipeNetworkIds
                    select ts.GetObject(networkId, OpenMode.ForRead)
                    as Network
                );
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
            if (DataSheet != null && DataSheet.IsReady())
                throw new DataSheetAlreadyExists();

            if (DataSheet != null && !DataSheet.IsReady())
            {
                DataSheet.Dispose();
                DataSheet = null;
            }
            try
            {
                DataSheet
                    = new PipeDataSheet(filePath, PipeDataSheet.PipeDataSheetName);

                // Test to see if it succeeded
                if (!DataSheet.IsReady())
                    DataSheet = null;
            }
            catch (FilenameNullException)
            {
                MessageBox.Show("No Filename was provided", "AutoSheet Error", MessageBoxButtons.OK);
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("Filename provided could not be found", "AutoSheet Error", MessageBoxButtons.OK);
            }
            catch (COMException e)
            {
                MessageBox.Show($"There was an error opening the file: {e}", "AutoSheet Error", MessageBoxButtons.OK);
                DataSheet?.Dispose();
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
            // exported so exit the function
            if (DataSheet == null) return;

            // Row number starts at 0 to reference the data at 0 in the arrays
            // The actual row data in the excel sheet should start at row 2
            var rowNumber = 0;
            var pipesIds = pipeNetwork.GetPipeIds();

            using (var ts = BootstrapApp.TransManager.StartTransaction())
            {
                // Arrays for the different attributes of a pipe
                var handleArray = new object[pipesIds.Count, 1];
                var fromArray = new object[pipesIds.Count, 1];
                var toArray = new object[pipesIds.Count, 1];
                var lengthArray = new object[pipesIds.Count, 1];
                var diameterArray = new object[pipesIds.Count, 1];

                foreach (ObjectId id in pipesIds)
                {
                    var pipe = ts.GetObject(id, OpenMode.ForRead) as Pipe;

                    // If we cannot access the pipe then skip it
                    if (pipe == null) continue;

                    handleArray[rowNumber, 0] = pipe.Handle.Value;

                    // Get start and End Structures and test to make sure that if there is none,
                    // that no errors are thrown
                    var startStructure = pipe.StartStructureId.IsNull
                        ? null
                        : ts.GetObject(pipe.StartStructureId, OpenMode.ForRead) as Structure;

                    var endStructure = pipe.EndStructureId.IsNull
                        ? null
                        : ts.GetObject(pipe.EndStructureId, OpenMode.ForRead) as Structure;

                    fromArray[rowNumber, 0] = startStructure == null ? "null" : startStructure.Name;
                    toArray[rowNumber, 0] = endStructure == null ? "null" : endStructure.Name;

                    lengthArray[rowNumber, 0] = pipe.Length2DCenterToCenter;
                    diameterArray[rowNumber, 0] = pipe.InnerDiameterOrWidth;

                    // Increment so that the next row can be filled out
                    rowNumber++;
                }

                DataSheet.GetRangeFromColumn(
                        DataSheet.HandleRange.Column, 2, pipesIds.Count + 1)
                        .Value2 = handleArray;

                DataSheet.GetRangeFromColumn(
                        DataSheet.FromRange.Column, 2, pipesIds.Count + 1)
                        .Value2 = fromArray;

                DataSheet.GetRangeFromColumn(
                        DataSheet.ToRange.Column, 2, pipesIds.Count + 1)
                        .Value2 = toArray;

                DataSheet.GetRangeFromColumn(
                        DataSheet.LengthRange.Column, 2, pipesIds.Count + 1)
                        .Value2 = lengthArray;

                DataSheet.GetRangeFromColumn(
                        DataSheet.InnerDiameterRange.Column, 2, pipesIds.Count + 1)
                        .Value2 = diameterArray;
            }
        }

        /// <summary>
        /// Imports Pipe Data into Excel
        /// </summary>
        /// <param name="network">The network that will be imported</param>
        public static void ImportPipeData(Network network)
        {
            // If the design sheet is null or not ready return
            if (DataSheet == null || !DataSheet.IsReady() || network == null)
                throw new ArgumentNullException();

            using (var ts = BootstrapApp.TransManager.StartTransaction())
            {
                BootstrapApp.ActiveDocument.LockDocument();
                var pipeIds = network.GetPipeIds();

                foreach (ObjectId pipeId in pipeIds)
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
                        pipe.StartPoint = new Point3d(pipe.StartPoint.X, pipe.StartPoint.Y, startInv + pipe.OuterDiameterOrWidth / 2 - pipe.WallThickness);
                        pipe.EndPoint = new Point3d(pipe.EndPoint.X, pipe.EndPoint.Y, endInv + pipe.OuterDiameterOrWidth / 2 - pipe.WallThickness);

                        // Disconnect and Reconnect the Start Structure
                        var startStructureId = pipe.StartStructureId;
                        pipe.Disconnect(ConnectorPositionType.Start);
                        pipe.ConnectToStructure(ConnectorPositionType.Start, startStructureId, true);

                        var startStructure = ts.GetObject(startStructureId, OpenMode.ForWrite) as Structure;
                        if (startStructure != null)
                        {
                            startStructure.ControlSumpBy = StructureControlSumpType.ByDepth;
                            startStructure.SumpDepth = 0.3;
                        }


                        // Disconnect and Reconnect the End Structure
                        var endStructureId = pipe.EndStructureId;
                        pipe.Disconnect(ConnectorPositionType.End);
                        pipe.ConnectToStructure(ConnectorPositionType.End, endStructureId, true);

                        var endStructure = ts.GetObject(endStructureId, OpenMode.ForWrite) as Structure;

                        if (endStructure != null)
                        {
                            endStructure.ControlSumpBy = StructureControlSumpType.ByDepth;
                            endStructure.SumpDepth = 0.3;
                        }

                        pipe.HoldOnResizeType = HoldOnResizeType.Crown;

                    }
                    catch (COMException exception)
                    {
                        // Notify the user and then abort the transaction
                        const string message =
                            "There was an error importing the data into AutoCAD, " +
                            "check to see if the pipe data table has the correct " +
                            "handles required for the VLOOKUP to function " +
                            "No objects have been updated " +
                            "\n\nExporting the Network should fix this issue";
                        MessageBox.Show($"{message}\n\n'{exception.Message}'", "AutoSheet Error", MessageBoxButtons.OK);
                        ts.Abort();
                        return;
                    }
                }
                BootstrapApp.ActiveDocument.Editor.Regen();

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

            if (DataSheet == null) return;
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

        private static bool ArePipesEqual(Pipe pipe1, Pipe pipe2)
        {
            if ( pipe1.Handle.Value != pipe2.Handle.Value )
                return false;

            return pipe1.StartPoint == pipe2.StartPoint && pipe1.EndPoint == pipe2.EndPoint;
        }

        #endregion
    }
}
