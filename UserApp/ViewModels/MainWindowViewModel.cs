using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace UserApp.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        private string _testString = "Adam";
        private Window _window;
        private bool _isVisible;

        public MainWindowViewModel(Window window)
        {
            _window = window;
            _window?.Show();
            _isVisible = true;

            testAction = () => { System.Diagnostics.Debug.WriteLine("It Worked"); };

            TestCommand = new CommandBase(testAction);
        }

        /// <summary>
        /// Returns true if the Window is Visible
        /// </summary>
        public bool IsVisable
        {
            get => _isVisible;

            private set
            {
                _isVisible = value;
                RaisePropertyChanged();
            }
        }

        public string TestString
        {
            get => _testString;

            set
            {
                _testString = value;
                RaisePropertyChanged();
            }
        }

        private Action testAction;

        public ICommand TestCommand { get; set; }
    }
}
