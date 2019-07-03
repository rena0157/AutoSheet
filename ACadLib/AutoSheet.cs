using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACadLib.Utilities;
using Autodesk.AutoCAD.Runtime;
using UserApp;
using UserApp.ViewModels;

namespace ACadLib
{
    public class AutoSheet
    {
        #region Private Fields

        private MainWindow _mainWindow;
        private MainWindowViewModel _mainWindowViewModel;
        private readonly ACadLogger _logger;

        #endregion
        
        #region Constructors

        /// <summary>
        /// Default Constructor
        /// </summary>
        public AutoSheet()
        {
            _logger = new ACadLogger(ACadLogger.LogLevel.Debug);
        }

        #endregion

        #region Public Properties

        

        #endregion

        #region AutoCAD Commands

        public void StartApplication()
        {
            if ( _mainWindow == null )
            {
                _mainWindow = new MainWindow();
                _mainWindowViewModel = _mainWindow?.DataContext as MainWindowViewModel;
                _mainWindow.Closed += MainWindowOnClosed;

                _logger.Log("Application Window Opened", ACadLogger.LogLevel.Debug);
            }
            else
                _logger.Log("Application Window Already Open", ACadLogger.LogLevel.Debug);
        }



        #endregion

        #region Private Methods

        private void MainWindowOnClosed(object sender, EventArgs e)
        {
            _mainWindow = null;
            _mainWindowViewModel = null;
            _logger.Log("Application Window Closed", ACadLogger.LogLevel.Debug);
        }

        #endregion
    }
}
