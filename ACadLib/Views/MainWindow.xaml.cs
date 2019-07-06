using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using UserApp.ViewModels;

namespace UserApp
{
    /// <summary>
    /// Interaction Logic for the Main Window
    /// </summary>
    public partial class AutoSheetMainWindow : Window
    {
        public AutoSheetMainWindow()
        {
            InitializeComponent();
            DataContext = new ViewModels.AutoSheetViewModel(this);
        }
    }
}
