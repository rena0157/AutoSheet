using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using ACadLib.Utilities;
using Autodesk.Civil.DatabaseServices;
using MessageBox = System.Windows.MessageBox;

namespace ACadLib.ViewModels
{
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

            BrowseCommand = new CommandBase(() => { CurrentPath = GetFileNameFileDialog(); });
            OpenDesignSheetCommand = new CommandBase((() => {AutoSheet.OpenDesignSheet(CurrentPath);}));

            // Run the Export Command
            ExportCommand = new CommandBase(() =>
            {
                if ( SelectedNetworkName != null )
                    AutoSheet.ExportPipeData(PipeNetworks[SelectedNetworkName]);
                else
                    MessageBox.Show("You must select a network", "AutoSheet Error", MessageBoxButton.OK);
            });

            // Run the import command
            ImportCommand = new CommandBase(() =>
            {
                if ( SelectedNetworkName != null )
                    AutoSheet.ImportPipeData(PipeNetworks[SelectedNetworkName]);
                else
                    MessageBox.Show("You must select a network", "AutoSheet Error", MessageBoxButton.OK);
            });
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
            get => _selectedNetwork?.Name ?? "None";
            set
            {
                if ( _pipeNetworks.ContainsKey(value) )
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
            };

            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : "";
        }

        #endregion
    }
}
