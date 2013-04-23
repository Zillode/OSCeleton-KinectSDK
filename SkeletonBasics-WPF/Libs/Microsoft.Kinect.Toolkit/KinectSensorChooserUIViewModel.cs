// --------------------------------------------------------------------------------------------------------------------
// <copyright file="KinectSensorChooserUIViewModel.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Microsoft.Kinect.Toolkit
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Data;
    using System.Windows.Input;
    using Microsoft.Kinect.Toolkit.Properties;

    /// <summary>
    /// View model for the KinectSensorChooser.
    /// </summary>
    public class KinectSensorChooserUIViewModel : DependencyObject
    {
        [SuppressMessage("StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "ReadOnlyDependencyProperty requires private static field to be initialized prior to the public static field")]
        private static readonly DependencyPropertyKey MessagePropertyKey = DependencyProperty.RegisterReadOnly(
            "Message", typeof(string), typeof(KinectSensorChooserUIViewModel), new PropertyMetadata(string.Empty));

        [SuppressMessage("StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "ReadOnlyDependencyProperty requires private static field to be initialized prior to the public static field")]
        private static readonly DependencyPropertyKey MoreInformationPropertyKey = DependencyProperty.RegisterReadOnly(
            "MoreInformation", typeof(string), typeof(KinectSensorChooserUIViewModel), new PropertyMetadata(string.Empty));

        [SuppressMessage("StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "ReadOnlyDependencyProperty requires private static field to be initialized prior to the public static field")]
        private static readonly DependencyPropertyKey MoreInformationUriPropertyKey =
            DependencyProperty.RegisterReadOnly(
                "MoreInformationUri", typeof(Uri), typeof(KinectSensorChooserUIViewModel), new PropertyMetadata(null));

        [SuppressMessage("StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "ReadOnlyDependencyProperty requires private static field to be initialized prior to the public static field")]
        private static readonly DependencyPropertyKey MoreInformationVisibilityPropertyKey =
            DependencyProperty.RegisterReadOnly(
                "MoreInformationVisibility",
                typeof(Visibility),
                typeof(KinectSensorChooserUIViewModel),
                new PropertyMetadata(Visibility.Collapsed));

        /// <summary>
        /// Set this to true when the application is listining for audio input from the sensor.
        /// UI will show a microphone icon.
        /// </summary>
        public static readonly DependencyProperty IsListeningProperty = DependencyProperty.Register(
            "IsListening", 
            typeof(bool), 
            typeof(KinectSensorChooserUIViewModel), 
            new PropertyMetadata(false, (o, args) => ((KinectSensorChooserUIViewModel)o).IsListeningChanged()));

        /// <summary>
        /// The KinectSensorChooser whose status we are displaying.
        /// </summary>
        public static readonly DependencyProperty KinectSensorChooserProperty = DependencyProperty.Register(
            "KinectSensorChooser", 
            typeof(KinectSensorChooser), 
            typeof(KinectSensorChooserUIViewModel), 
            new PropertyMetadata(
                null, 
                (o, args) => ((KinectSensorChooserUIViewModel)o).OnKinectKinectSensorChooserChanged((KinectSensorChooser)args.NewValue)));

        /// <summary>
        /// The current ChooserStatus of our KinectSensorChooser
        /// </summary>
        public static readonly DependencyProperty StatusProperty = DependencyProperty.Register(
            "Status", 
            typeof(ChooserStatus), 
            typeof(KinectSensorChooserUIViewModel), 
            new PropertyMetadata(ChooserStatus.None, (o, args) => ((KinectSensorChooserUIViewModel)o).OnStatusChanged()));

        /// <summary>
        /// The state we want the VisualStateManager in the UI to be in.
        /// </summary>
        public static readonly DependencyProperty VisualStateProperty = DependencyProperty.Register(
            "VisualState", typeof(string), typeof(KinectSensorChooserUIViewModel), new PropertyMetadata(null));

        /// <summary>
        /// The short message to display.
        /// </summary>
        public static readonly DependencyProperty MessageProperty = MessagePropertyKey.DependencyProperty;

        /// <summary>
        /// The more descriptive message to display.
        /// </summary>
        public static readonly DependencyProperty MoreInformationProperty = MoreInformationPropertyKey.DependencyProperty;

        /// <summary>
        /// Uri for more information on the state we are in.
        /// </summary>
        public static readonly DependencyProperty MoreInformationUriProperty = MoreInformationUriPropertyKey.DependencyProperty;

        /// <summary>
        /// The visibility we want for the MoreInformation Uri.
        /// </summary>
        public static readonly DependencyProperty MoreInformationVisibilityProperty = MoreInformationVisibilityPropertyKey.DependencyProperty;

        private RelayCommand retryCommand;

        /// <summary>
        /// Set this to true when the application is listining for audio input from the sensor.
        /// UI will show a microphone icon.
        /// </summary>
        public bool IsListening
        {
            get
            {
                return (bool)this.GetValue(IsListeningProperty);
            }

            set
            {
                this.SetValue(IsListeningProperty, value);
            }
        }

        /// <summary>
        /// The KinectSensorChooser whose status we are displaying.
        /// </summary>
        public KinectSensorChooser KinectSensorChooser
        {
            get
            {
                return (KinectSensorChooser)this.GetValue(KinectSensorChooserProperty);
            }

            set
            {
                this.SetValue(KinectSensorChooserProperty, value);
            }
        }

        /// <summary>
        /// The short message to display.
        /// </summary>
        public string Message
        {
            get
            {
                return (string)this.GetValue(MessageProperty);
            }

            private set
            {
                this.SetValue(MessagePropertyKey, value);
            }
        }

        /// <summary>
        /// The more descriptive message to display.
        /// </summary>
        public string MoreInformation
        {
            get
            {
                return (string)this.GetValue(MoreInformationProperty);
            }

            private set
            {
                this.SetValue(MoreInformationPropertyKey, value);
            }
        }

        /// <summary>
        /// Uri for more information on the state we are in.
        /// </summary>
        public Uri MoreInformationUri
        {
            get
            {
                return (Uri)this.GetValue(MoreInformationUriProperty);
            }

            private set
            {
                this.SetValue(MoreInformationUriPropertyKey, value);
            }
        }

        /// <summary>
        /// The visibility we want for the MoreInformation Uri.
        /// </summary>
        public Visibility MoreInformationVisibility
        {
            get
            {
                return (Visibility)this.GetValue(MoreInformationVisibilityProperty);
            }

            private set
            {
                this.SetValue(MoreInformationVisibilityPropertyKey, value);
            }
        }

        /// <summary>
        /// Command to retry getting a sensor.
        /// </summary>
        public ICommand RetryCommand
        {
            get
            {
                if (this.retryCommand == null)
                {
                    this.retryCommand = new RelayCommand(this.Retry, this.CanRetry);
                }

                return this.retryCommand;
            }
        }

        /// <summary>
        /// The current ChooserStatus of our KinectSensorChooser
        /// </summary>
        public ChooserStatus Status
        {
            get
            {
                return (ChooserStatus)this.GetValue(StatusProperty);
            }

            set
            {
                this.SetValue(StatusProperty, value);
            }
        }

        /// <summary>
        /// The state we want the VisualStateManager in the UI to be in.
        /// </summary>
        public string VisualState
        {
            get
            {
                return (string)this.GetValue(VisualStateProperty);
            }

            set
            {
                this.SetValue(VisualStateProperty, value);
            }
        }

        /// <summary>
        /// Determines if the retry command is available
        /// </summary>
        /// <returns>true if retry is valid, false otherwise</returns>
        private bool CanRetry()
        {
            // You can retry if another app was using the sensor the last
            // time we tried.  You can also retry if the only problem was
            // that there were no sensors available since this app may
            // have released one.
            return (0 != (this.Status & ChooserStatus.SensorConflict)) || (this.Status == ChooserStatus.NoAvailableSensors);
        }

        private void IsListeningChanged()
        {
            this.UpdateState();
        }

        private void OnKinectKinectSensorChooserChanged(KinectSensorChooser newValue)
        {
            if (newValue != null)
            {
                var statusBinding = new Binding("Status") { Source = newValue };
                BindingOperations.SetBinding(this, StatusProperty, statusBinding);
            }
            else
            {
                BindingOperations.ClearBinding(this, StatusProperty);
            }
        }

        private void OnStatusChanged()
        {
            this.UpdateState();
        }

        private void Retry()
        {
            if (this.KinectSensorChooser != null)
            {
                this.KinectSensorChooser.TryResolveConflict();
            }
        }

        private void UpdateState()
        {
            string newVisualState;
            string message;
            string moreInfo = null;
            Uri moreInfoUri = null;

            if ((this.Status & ChooserStatus.SensorStarted) != 0)
            {
                if (this.IsListening)
                {
                    newVisualState = "AllSetListening";
                    message = Resources.MessageAllSetListening;
                }
                else
                {
                    newVisualState = "AllSetNotListening";
                    message = Resources.MessageAllSet;
                }
            }
            else if ((this.Status & ChooserStatus.SensorInitializing) != 0)
            {
                newVisualState = "Initializing";
                message = Resources.MessageInitializing;
            }
            else if ((this.Status & ChooserStatus.SensorConflict) != 0)
            {
                newVisualState = "Error";
                message = Resources.MessageConflict;
                moreInfo = Resources.MoreInformationConflict;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239812");
            }
            else if ((this.Status & ChooserStatus.SensorNotGenuine) != 0)
            {
                newVisualState = "Error";
                message = Resources.MessageNotGenuine;
                moreInfo = Resources.MoreInformationNotGenuine;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239813");
            }
            else if ((this.Status & ChooserStatus.SensorNotSupported) != 0)
            {
                newVisualState = "Error";
                message = Resources.MessageNotSupported;
                moreInfo = Resources.MoreInformationNotSupported;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239814");
            }
            else if ((this.Status & ChooserStatus.SensorError) != 0)
            {
                newVisualState = "Error";
                message = Resources.MessageError;
                moreInfo = Resources.MoreInformationError;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239817");
            }
            else if ((this.Status & ChooserStatus.SensorInsufficientBandwidth) != 0)
            {
                newVisualState = "Error";
                message = Resources.MessageInsufficientBandwidth;
                moreInfo = Resources.MoreInformationInsufficientBandwidth;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239818");
            }
            else if ((this.Status & ChooserStatus.SensorNotPowered) != 0)
            {
                newVisualState = "Error";
                message = Resources.MessageNotPowered;
                moreInfo = Resources.MoreInformationNotPowered;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239819");
            }
            else if ((this.Status & ChooserStatus.NoAvailableSensors) != 0)
            {
                newVisualState = "NoAvailableSensors";
                message = Resources.MessageNoAvailableSensors;
                moreInfo = Resources.MoreInformationNoAvailableSensors;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239815");
            }
            else
            {
                newVisualState = "Stopped";
                message = Resources.MessageNoAvailableSensors;
                moreInfo = Resources.MoreInformationNoAvailableSensors;
                moreInfoUri = new Uri("http://go.microsoft.com/fwlink/?LinkID=239815");
            }

            this.Message = message;
            this.MoreInformation = moreInfo;
            this.MoreInformationUri = moreInfoUri;
            this.MoreInformationVisibility = moreInfoUri == null ? Visibility.Collapsed : Visibility.Visible;
            if (this.retryCommand != null)
            {
                this.retryCommand.InvokeCanExecuteChanged();
            }

            this.VisualState = newVisualState;
        }
    }
}