using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;

using System.Diagnostics;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using TransactionManager = Autodesk.AutoCAD.DatabaseServices.TransactionManager;

using ACadLib.Utilities;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ACadLib
{
    /// <summary>
    /// Main bootstrapping application for AutoCAD
    /// </summary>
    // ReSharper disable once ClassNeverInstantiated.Global
    public class BootstrapApp : IExtensionApplication
    {
        /// <summary>
        /// The Active Document for core AutoCAD Functions
        /// </summary>
        public static Document ActiveDocument 
            => Application.DocumentManager.MdiActiveDocument;

        /// <summary>
        /// The Transaction Manager for AutoCAD Entities
        /// </summary>
        public static TransactionManager TransManager
            => ActiveDocument.TransactionManager;

        /// <summary>
        /// The Civil Document For Civil Entities
        /// </summary>
        public static CivilDocument CivilDoc
            => CivilApplication.ActiveDocument;

        /// <summary>
        /// Initialization Function for the application
        /// </summary>
        public void Initialize()
        {
            ACadLogger.Log("Application Loaded");
        }

        /// <summary>
        /// Termination function for the application
        /// </summary>
        public void Terminate()
        {

        }

        #region Command Methods

        /// <summary>
        /// Command That will run AutoSheet and open the
        /// AutoSheet main Window
        /// </summary>
        [CommandMethod("AUTOSHEET")]
        public void AutoSheetCommand()
        {
            ACadLogger.Log("Starting AutoSheet Command");
            AutoSheet.StartApplication();
        }

        #endregion

    }
}
