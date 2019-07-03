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
using UserApp;
using UserApp.ViewModels;

using ACadLib.Utilities;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ACadLib
{
    /// <summary>
    /// Main bootstrapping application for AutoCAD
    /// </summary>
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

        public static AutoSheet AutoSheetApp { get; private set; }

        /// <summary>
        /// Initialization Function for the application
        /// </summary>
        public void Initialize()
        {
            ACadLogger.Log("Application Loaded");
            AutoSheetApp = new AutoSheet();
        }

        /// <summary>
        /// Termination function for the application
        /// </summary>
        public void Terminate()
        {

        }

        [CommandMethod("AUTOSHEET")]
        public void AutoSheetCommand()
        {
            ACadLogger.Log("Starting AutoSheet Command");
            AutoSheetApp.StartApplication();
        }

    }
}
