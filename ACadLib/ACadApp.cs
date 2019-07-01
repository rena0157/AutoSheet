using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;

using System.Diagnostics;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using TransactionManager = Autodesk.AutoCAD.DatabaseServices.TransactionManager;
using UserApp;

namespace ACadLib
{
    /// <summary>
    /// Application that will be loaded into AutoCAD
    /// </summary>
    public class ACadApp : IExtensionApplication
    {

        public Editor ACadEditor { get; private set; }

        public CivilDocument CDocument { get; private set; }

        public Document ActiveDocument { get; private set; }

        public List<Network> PipeNetworks { get; private set; }

        private TransactionManager _transactionManager;

        private MainWindow _mainWindow;

        /// <summary>
        /// Initialization Function for the application
        /// </summary>
        public void Initialize()
        {
            // Get the Active Document
            ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            // Get the active editor
            // ReSharper disable once AccessToStaticMemberViaDerivedType
            ACadEditor = ActiveDocument.Editor;

            ACadEditor.WriteMessage("Let's Try This....");

            // Get the active document
            CDocument = CivilApplication.ActiveDocument;

            _transactionManager = ActiveDocument.TransactionManager;

            var ids = GetPipeNetworksIds();

        }

        /// <summary>
        /// Termination function for the application
        /// </summary>
        public void Terminate()
        {

        }

        [CommandMethod("AUTOSHEET")]
        public void StartApplication()
        {
            if (_mainWindow == null) _mainWindow = new MainWindow();
        }

        private ObjectIdCollection GetPipeNetworksIds()
        {
            // Check that there is a pipe network to parse
            // and if there is then return the network ids
            if ( CDocument.GetPipeNetworkIds() != null )
                return CDocument.GetPipeNetworkIds();

            ACadEditor.WriteMessage("\nThere are no pipe networks to parse.");
            return new ObjectIdCollection();
        }
    }
}
