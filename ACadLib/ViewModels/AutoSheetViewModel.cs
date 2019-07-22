// AutoSheetViewModel.cs
// By: Adam Renaud
// Created: 2019-07-21

using ACadLib.Exceptions;
using ACadLib.Utilities;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using MessageBox = System.Windows.MessageBox;

namespace ACadLib.ViewModels
{
    /// <summary>
    /// The view model for the AutoSheet Window
    /// </summary>
    public class AutoSheetViewModel : ViewModelBase
    {
        #region Private Fields

        /// <summary>
        /// The Window That is passed to this class
        /// during construction
        /// </summary>
        private readonly Window _window;

        /// <summary>
        /// A dictionary of all Pipe networks
        /// </summary>
        private Dictionary<string, Network> _pipeNetworks;

        /// <summary>
        /// The Current Selected Network
        /// </summary>
        private Network _selectedNetwork;

        /// <summary>
        /// The Path to the current Excel File
        /// </summary>
        private string _currentPath;

        #endregion

        #region Constructors

        /// <summary>
        /// Default Constructor
        /// </summary>
        /// <param name="window">The Window</param>
        public AutoSheetViewModel(Window window)
        {
            _window = window;
            _window?.Show();

            _pipeNetworks = new Dictionary<string, Network>();

            // Set the commands and the command methods
            BrowseCommand = new CommandBase(BrowseCommandMethod);
            OpenDesignSheetCommand = new CommandBase(OpenDesignSheetCommandMethod);
            ExportCommand = new CommandBase(ExportCommandMethod);
            ImportCommand = new CommandBase(ImportCommandMethod);
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// A dictionary of all the Pipe Networks
        /// </summary>
        public Dictionary<string, Network> PipeNetworks
        {
            get
            {
                _pipeNetworks = new Dictionary<string, Network>();
                foreach (var network in AutoSheet.GetPipeNetworks())
                    _pipeNetworks.Add(network.Name, network);

                return _pipeNetworks;
            }
            set
            {
                _pipeNetworks = value;
                RaisePropertyChanged();
            }
        }

        /// <summary>
        /// The Name of the Current Network
        /// </summary>
        public string SelectedNetworkName
        {
            get => _selectedNetwork?.Name;
            set
            {
                if (_pipeNetworks.ContainsKey(value))
                    _selectedNetwork = _pipeNetworks[value];
                RaisePropertyChanged();
                ACadLogger.Log($"The Selected Network is now: {value}");
            }
        }

        /// <summary>
        /// The Path to the current Excel File
        /// </summary>
        public string CurrentPath
        {
            get => _currentPath;
            set
            {
                _currentPath = value;
                RaisePropertyChanged();
            }
        }

        #endregion

        #region Commands

        /// <summary>
        /// Opens the File Dialog, gets a filename from the user
        /// and sets that values to <see cref="CurrentPath"/>
        /// </summary>
        public ICommand BrowseCommand { get; set; }

        /// <summary>
        /// Command That Opens the Current Selected Design Sheet
        /// </summary>
        public ICommand OpenDesignSheetCommand { get; set; }

        /// <summary>
        /// Export Data to the Design Sheet Command
        /// </summary>
        public ICommand ExportCommand { get; set; }

        /// <summary>
        /// Import Data from the Design Sheet Command
        /// </summary>
        public ICommand ImportCommand { get; set; }

        #endregion

        #region Command Methods

        /// <summary>
        /// A Command Method that runs the import command
        /// </summary>
        private void ImportCommandMethod()
        {
            // If there is no design sheet then we cannot run the command
            if (AutoSheet.DataSheet == null || !AutoSheet.DataSheet.IsReady())
            {
                MessageBox.Show("You must open a design sheet first", "AutoSheet Error", MessageBoxButton.OK);
                return;
            }

            if (SelectedNetworkName == null)
            {
                MessageBox.Show("You Must select a network first", "AutoSheet Error", MessageBoxButton.OK);
                return;
            }

            try
            {
                AutoSheet.ImportPipeData(PipeNetworks[SelectedNetworkName]);
            }
            catch (ArgumentNullException)
            {
                MessageBox.Show("Either the Design Sheet is not open or the selected Network does not exist",
                    "AutoSheet Error", MessageBoxButton.OK);
            }

        }

        /// <summary>
        /// Export Command Method, runs the command method for exporting
        /// Data from Excel to AutoCAD
        /// </summary>
        private void ExportCommandMethod()
        {
            // If there is no design sheet then we cannot run the command
            if (AutoSheet.DataSheet == null || !AutoSheet.DataSheet.IsReady())
            {
                MessageBox.Show("You must open a design sheet first", "AutoSheet Error", MessageBoxButton.OK);
                return;
            }

            if (SelectedNetworkName == null)
            {
                MessageBox.Show("You must select a network first", "AutoSheet Error", MessageBoxButton.OK);
                return;
            }

            AutoSheet.ExportPipeData(PipeNetworks[SelectedNetworkName]);
        }

        /// <summary>
        /// Command Method for the Open Design Sheet Command
        /// </summary>
        private void OpenDesignSheetCommandMethod()
        {
            try
            {
                AutoSheet.OpenDesignSheet(CurrentPath);
            }
            catch (DataSheetAlreadyExists)
            {
                MessageBox.Show("There is already a design sheet that is open and running", "AutoSheet Error",
                    MessageBoxButton.OK);
            }
        }

        /// <summary>
        /// Command Method for the Browse command
        /// </summary>
        private void BrowseCommandMethod()
        {
            CurrentPath = GetFileNameFileDialog();
            OpenDesignSheetCommandMethod();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Opens the <see cref="OpenFileDialog"/> and returns a string
        /// that is the file path
        /// </summary>
        /// <returns>A Path to the file selected</returns>
        private static string GetFileNameFileDialog()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select a Design Sheet",
                CheckFileExists = true,
                Filter = "Excel Macro Files (*.xlsm) | *.xlsm"
            };

            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : "";
        }

        #endregion

    }
}
