using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ACadLib.ViewModels
{
    /// <summary>
    /// A Basic Command the executes an action
    /// </summary>
    public class CommandBase : ICommand
    {
        /// <summary>
        /// The action to run
        /// </summary>
        private readonly Action _action;

        /// <summary>
        /// Construction of the Command from an action that is
        /// to be completed
        /// </summary>
        /// <param name="action">The action that is to be executed</param>
        public CommandBase(Action action) => _action = action;
        
        /// <summary>
        /// The action can always execute here
        /// </summary>
        public bool CanExecute(object parameter) => true;

        /// <summary>
        /// Execute the action
        /// </summary>
        public void Execute(object parameter) => _action();
        
        /// <summary>
        /// The event that says if the Can Execute changed.
        /// </summary>
        public event EventHandler CanExecuteChanged = (sender, e) => { };
    }
}
