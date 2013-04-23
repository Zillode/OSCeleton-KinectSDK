// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RelayCommand.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Microsoft.Kinect.Toolkit
{
    using System;
    using System.Globalization;
    using System.Windows.Input;
    using Microsoft.Kinect.Toolkit.Properties;

    /// <summary>
    /// Command that executes a delegate that takes no parameters.
    /// </summary>
    public class RelayCommand : ICommand
    {
        /// <summary>
        /// Delegate to be executed 
        /// </summary>
        private Action executeDelegate;

        /// <summary>
        /// Predicate determining whether this command can currently execute
        /// </summary>
        private Func<bool> canExecuteDelegate;

        private EventHandler canExecuteEventhandler;

        /// <summary>
        /// Initializes a new instance of the RelayCommand class with the provided delegate and predicate
        /// </summary>
        /// <param name="executeDelegate">Delegate to be executed</param>
        /// <param name="canExecuteDelegate">Predicate determining whether this command can currently execute</param>
        public RelayCommand(Action executeDelegate, Func<bool> canExecuteDelegate)
        {
            if (null == executeDelegate)
            {
                throw new ArgumentNullException("executeDelegate");
            }

            this.canExecuteDelegate = canExecuteDelegate;
            this.executeDelegate = executeDelegate;
        }

        /// <summary>
        /// Initializes a new instance of the RelayCommand class with the provided delegate
        /// </summary>
        /// <param name="executeDelegate">Delegate to be executed</param>
        public RelayCommand(Action executeDelegate)
            : this(executeDelegate, null)
        {
        }

        /// <summary>
        /// Event signaling that the possibility of this command executing has changed
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add
            {
                this.canExecuteEventhandler += value;
                CommandManager.RequerySuggested += value;
            }

            remove
            {
                this.canExecuteEventhandler -= value;
                CommandManager.RequerySuggested -= value;
            }
        }

        /// <summary>
        /// Evaluates whether the command can currently execute
        /// </summary>
        /// <param name="parameter">ICommand required parameter that is ignored</param>
        /// <returns>True if the command can currently execute, false otherwise</returns>
        public bool CanExecute(object parameter)
        {
            if (null == this.canExecuteDelegate)
            {
                return true;
            }

            return this.canExecuteDelegate.Invoke();
        }

        /// <summary>
        /// Executes the associated delegate
        /// </summary>
        /// <param name="parameter">ICommand required parameter that is ignored</param>
        public void Execute(object parameter)
        {
            this.executeDelegate.Invoke();
        }

        /// <summary>
        /// Raises the CanExecuteChanged event to signal that the possibility of execution has changed
        /// </summary>
        public void InvokeCanExecuteChanged()
        {
            if (null != canExecuteDelegate)
            {
                EventHandler handler = this.canExecuteEventhandler;
                if (null != handler)
                {
                    handler(this, EventArgs.Empty);
                }
            }
        }
    }

    public class RelayCommand<T> : ICommand where T : class
    {
        /// <summary>
        /// Delegate to be executed 
        /// </summary>
        private Action<T> executeDelegate;

        /// <summary>
        /// Predicate determining whether this command can currently execute
        /// </summary>
        private Predicate<T> canExecuteDelegate;

        private EventHandler canExecuteEventhandler;

        /// <summary>
        /// Initializes a new instance of the RelayCommand class with the provided delegate and predicate
        /// </summary>
        /// <param name="executeDelegate">Delegate to be executed</param>
        /// <param name="canExecuteDelegate">Predicate determining whether this command can currently execute</param>
        public RelayCommand(Action<T> executeDelegate, Predicate<T> canExecuteDelegate)
        {
            if (null == executeDelegate)
            {
                throw new ArgumentNullException("executeDelegate");
            }

            this.canExecuteDelegate = canExecuteDelegate;
            this.executeDelegate = executeDelegate;
        }

        /// <summary>
        /// Initializes a new instance of the RelayCommand class with the provided delegate
        /// </summary>
        /// <param name="executeDelegate">Delegate to be executed</param>
        public RelayCommand(Action<T> executeDelegate)
            : this(executeDelegate, null)
        {
        }

        /// <summary>
        /// Event signaling that the possibility of this command executing has changed
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add
            {
                this.canExecuteEventhandler += value;
                CommandManager.RequerySuggested += value;
            }

            remove
            {
                this.canExecuteEventhandler -= value;
                CommandManager.RequerySuggested -= value;
            }
        }

        /// <summary>
        /// Evaluates whether the command can currently execute
        /// </summary>
        /// <param name="parameter">Context of type T used for evaluating the current possibility of execution</param>
        /// <returns>True if the command can currently execute, false otherwise</returns>
        public bool CanExecute(object parameter)
        {
            if (null == parameter)
            {
                throw new ArgumentNullException("parameter");
            }

            if (null == this.canExecuteDelegate)
            {
                return true;
            }

            T castParameter = parameter as T;
            if (null == castParameter)
            {
                throw new InvalidCastException(string.Format(CultureInfo.InvariantCulture, Resources.DelegateCommandCastException, parameter.GetType().FullName, typeof(T).FullName));
            }

            return this.canExecuteDelegate.Invoke(castParameter);
        }

        /// <summary>
        /// Executes the associated delegate
        /// </summary>
        /// <param name="parameter">Parameter of type T passed to the associated delegate</param>
        public void Execute(object parameter)
        {
            if (null == parameter)
            {
                throw new ArgumentNullException("parameter");
            }

            T castParameter = parameter as T;
            if (null == castParameter)
            {
                throw new InvalidCastException(string.Format(CultureInfo.InvariantCulture, Resources.DelegateCommandCastException, parameter.GetType().FullName, typeof(T).FullName));
            }

            this.executeDelegate.Invoke(castParameter);
        }

        /// <summary>
        /// Raises the CanExecuteChanged event to signal that the possibility of execution has changed
        /// </summary>
        public void InvokeCanExecuteChanged()
        {
            if (null != canExecuteDelegate)
            {
                EventHandler handler = this.canExecuteEventhandler;
                if (null != handler)
                {
                    handler(this, EventArgs.Empty);
                }
            }
        }
    }
}