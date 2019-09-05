// BootstrapApp.cs
// By: Adam Renaud
// Created: July 2019

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using System.Diagnostics;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using TransactionManager = Autodesk.AutoCAD.DatabaseServices.TransactionManager;

namespace ACadLib
{
    /// <summary>
    /// Main bootstrapping application for AutoCAD
    /// </summary>
    public class BootstrapApp : IExtensionApplication
    {
        #region Static AutoCAD Properties

        /// <summary>
        /// The Active <see cref="Document"/> for core AutoCAD Functions
        /// </summary>
        public static Document ActiveDocument
            => Application.DocumentManager.MdiActiveDocument;

        /// <summary>
        /// The Transaction Manager for AutoCAD Entities that
        /// are a part of this <see cref="Document"/>
        /// </summary>
        public static TransactionManager TransManager
            => ActiveDocument.TransactionManager;

        /// <summary>
        /// The Civil Document For Civil Entities
        /// </summary>
        public static CivilDocument CivilDoc
            => CivilApplication.ActiveDocument;

        #endregion

        #region IExtension Methods

        /// <summary>
        /// Initialization Function for the application
        /// </summary>
        public void Initialize()
        {
            Debug.WriteLine("Application Initialized");
        }

        /// <summary>
        /// Termination function for the application
        /// </summary>
        public void Terminate()
        {
            if (AutoSheet.DataSheet != null)
                AutoSheet.DataSheet.Dispose();

            Debug.WriteLine("Application Terminated");
        }

        #endregion

        #region Command Methods

        /// <summary>
        /// Command That will run AutoSheet and open the
        /// AutoSheet main Window
        /// </summary>
        [CommandMethod("AUTOSHEET")]
        public void AutoSheetCommand()
        {
            Debug.WriteLine("Starting Application");
            AutoSheet.StartApplication();
        }

        #endregion
    }
}
