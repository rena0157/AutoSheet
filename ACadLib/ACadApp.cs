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

namespace ACadLib
{
    /// <summary>
    /// Application that will be loaded into AutoCAD
    /// </summary>
    public class ACadApp : IExtensionApplication
    {

        public CivilDocument CDocument { get; private set; }

        public Document ActiveDocument { get; private set; }


        private TransactionManager _transactionManager;
        private MainWindow _mainWindow;
        private MainWindowViewModel _mainWindowVm;
        private ACadLogger _logger;

        /// <summary>
        /// Initialization Function for the application
        /// </summary>
        public void Initialize()
        {
            // Get the Active Document
            // ReSharper disable once AccessToStaticMemberViaDerivedType
            ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            // Get the active document
            CDocument = CivilApplication.ActiveDocument;

            // Set the Transaction Manager
            _transactionManager = ActiveDocument.TransactionManager;

            // Set the logger
            _logger = new ACadLogger(ActiveDocument);

            _logger.Log("Application Loaded");
        }

        /// <summary>
        /// Termination function for the application
        /// </summary>
        public void Terminate()
        {
            _logger.Log("Application Closing");
        }

        [CommandMethod("AUTOSHEET")]
        public void StartApplication()
        {
            if ( _mainWindow == null )
            {
                _mainWindow = new MainWindow();
                _mainWindowVm = _mainWindow.DataContext as MainWindowViewModel;
                _mainWindow.Closed += MainWindowOnClosing;

                _logger.Log("Main Window Created");
            }
            else
                ActiveDocument.Editor.WriteMessage("Window Already Open");
        }

        /// <summary>
        /// When the window is closing run this
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainWindowOnClosing(object sender, EventArgs e)
        {
            _mainWindow = null;
            _mainWindowVm = null;
        }

        private ObjectIdCollection GetPipeNetworksIds()
        {
            // Check that there is a pipe network to parse
            // and if there is then return the network ids
            if ( CDocument.GetPipeNetworkIds() != null )
                return CDocument.GetPipeNetworkIds();

            return new ObjectIdCollection();
        }
    }
}
