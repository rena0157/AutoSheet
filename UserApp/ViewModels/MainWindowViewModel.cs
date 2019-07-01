using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace UserApp.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        private string _testString = "Adam";

        public MainWindowViewModel()
        {
            testAction = () => { System.Diagnostics.Debug.WriteLine("It Worked"); };

            TestCommand = new CommandBase(testAction);
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
