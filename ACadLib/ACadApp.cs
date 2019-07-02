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
    /// Main Application that is loaded into AutoCAD
    /// </summary>
    public class ACadApp : IExtensionApplication
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


        private MainWindow _mainWindow;
        private MainWindowViewModel _mainWindowVm;
        private ACadLogger _logger;

        /// <summary>
        /// Initialization Function for the application
        /// </summary>
        public void Initialize()
        {
            // Set the logger
            _logger = new ACadLogger(ACadLogger.LogLevel.Debug);

            ACadLogger.Log("Application Loaded");
        }

        /// <summary>
        /// Termination function for the application
        /// </summary>
        public void Terminate()
        {
            _logger.Log("Application Closing", ACadLogger.LogLevel.Debug);
        }

        /// <summary>
        /// Start the application via the Command Line
        /// </summary>
        [CommandMethod("AUTOSHEET")]
        public void StartApplication()
        {
            if ( _mainWindow == null )
            {
                _mainWindow = new MainWindow();
                _mainWindowVm = _mainWindow.DataContext as MainWindowViewModel;
                _mainWindow.Closed += MainWindowOnClosing;

                _logger.Log("Main Window Created", ACadLogger.LogLevel.Debug);
            }
            else
                _logger.Log("Window Already Open", ACadLogger.LogLevel.Debug);
        }

        /// <summary>
        /// Clean up resources when the main window closes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainWindowOnClosing(object sender, EventArgs e)
        {
            _mainWindow = null;
            _mainWindowVm = null;
            _logger.Log("Main Window Closed", ACadLogger.LogLevel.Debug);
        }
    }
}
