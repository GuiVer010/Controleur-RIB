using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Contrôleur_RIB
{
    /// <summary>
    /// Ceci est une classe custom qui nous sert pour faire le binding entre les page en xml et le code qu'il y a derrière
    /// </summary>
    public class ViewModelBase : INotifyPropertyChanged
    {

        #region INotifyPropertyChanged Members  

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (propertyName != null)
                this.VerifyPropertyName(propertyName);
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                var e = new PropertyChangedEventArgs(propertyName);
                handler(this, e);
            }
        }

        private bool ThrowOnInvalidPropertyName = true;
        [Conditional("DEBUG")]
        [DebuggerStepThrough]
        public void VerifyPropertyName(string propertyName)
        {
            // Verify that the property name matches a real, 
            // public, instance property on this object. 
            if (TypeDescriptor.GetProperties(this)[propertyName] == null)
            {
                string msg = "Invalid property name: " + propertyName;
                if (this.ThrowOnInvalidPropertyName)
                    throw new Exception(msg);
                else
                    Debug.Fail(msg);
            }
        }
        #endregion


        #region Command
        public class Command : ICommand
        {
            Action<object> _TargetExecuteMethodParam;
            Action _TargetExecuteMethod;
            Func<bool> _TargetCanExecuteMethod;

            public Command(Action executeMethod)
            {
                _TargetExecuteMethod = executeMethod;
            }

            public Command(Action executeMethod, Func<bool> canExecuteMethod)
            {
                _TargetExecuteMethod = executeMethod;
                _TargetCanExecuteMethod = canExecuteMethod;
            }

            public Command(Action<object> executeMethod)
            {
                _TargetExecuteMethodParam = executeMethod;
            }

            public Command(Action<object> executeMethod, Func<bool> canExecuteMethod)
            {
                _TargetExecuteMethodParam = executeMethod;
                _TargetCanExecuteMethod = canExecuteMethod;
            }

            public void RaiseCanExecuteChanged()
            {
                CanExecuteChanged(this, EventArgs.Empty);
            }

            bool ICommand.CanExecute(object parameter)
            {

                if (_TargetCanExecuteMethod != null)
                {
                    return _TargetCanExecuteMethod();
                }

                if (_TargetExecuteMethodParam != null)
                {
                    return true;
                }

                if (_TargetExecuteMethod != null)
                {
                    return true;
                }

                return false;
            }

            // Beware - should use weak references if command instance lifetime 
            //  is longer than lifetime of UI objects that get hooked up to command

            // Prism commands solve this in their implementation 
            public event EventHandler CanExecuteChanged = delegate { };

            void ICommand.Execute(object parameter)
            {
                if (_TargetExecuteMethodParam != null)
                {
                    _TargetExecuteMethodParam(parameter);
                }
                else if (_TargetExecuteMethod != null)
                {
                    _TargetExecuteMethod();
                }
            }
        }
        #endregion

    }
}
