// --------------------------------------------------------------------------------------------------------------------
// <copyright file="KinectSensorChooserUI.xaml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Microsoft.Kinect.Toolkit
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Navigation;
    using System.Windows.Threading;

    /// <summary>
    /// Interaction logic for KinectSensorChooserUI.xaml
    /// </summary>
    public partial class KinectSensorChooserUI : UserControl
    {
        /// <summary>
        /// Set this to true when the application is listining for audio input from the sensor.
        /// UI will show a microphone icon.
        /// </summary>
        public static readonly DependencyProperty IsListeningProperty = DependencyProperty.Register(
            "IsListening", typeof(bool), typeof(KinectSensorChooserUI), new PropertyMetadata(false));

        /// <summary>
        /// The KinectSensorChooser whose status we are displaying.
        /// </summary>
        public static readonly DependencyProperty KinectSensorChooserProperty = DependencyProperty.Register(
            "KinectSensorChooser", typeof(KinectSensorChooser), typeof(KinectSensorChooserUI), new PropertyMetadata(null));

        /// <summary>
        /// Used internally to transfer the visual state from the view model to the
        /// VisualStateManager defined in XAML.
        /// </summary>
        public static readonly DependencyProperty VisualStateProperty = DependencyProperty.Register(
            "VisualState", 
            typeof(string), 
            typeof(KinectSensorChooserUI), 
            new PropertyMetadata(null, (o, args) => ((KinectSensorChooserUI)o).OnVisualstateChanged((string)args.NewValue)));

        /// <summary>
        /// Timer used to check if the mouse is still in the popup when we activated
        /// the popup with a mouse hover.  We do this because the mouse may not be
        /// in the popup when it comes up and we will never get a mouse leave event.
        /// </summary>
        private readonly DispatcherTimer popupCloseCheck;

        private bool suppressPopupOnFocus;

        private Window parentWindow;

        /// <summary>
        /// Initializes a new instance of the KinectSensorChooserUI class
        /// </summary>
        public KinectSensorChooserUI()
        {
            Loaded += OnLoaded;
            Unloaded += OnUnloaded;

            this.InitializeComponent();
            this.popupCloseCheck = new DispatcherTimer(
                TimeSpan.FromMilliseconds(1000), DispatcherPriority.Normal, this.OnPopupCloseCheckFired, this.Dispatcher);

            var viewModel = new KinectSensorChooserUIViewModel();
            this.layoutRoot.DataContext = viewModel;

            var visualStateBinding = new Binding("VisualState") { Source = viewModel };
            SetBinding(VisualStateProperty, visualStateBinding);

            var sensorChooserBinding = new Binding("KinectSensorChooser") { Source = this };
            BindingOperations.SetBinding(viewModel, KinectSensorChooserUIViewModel.KinectSensorChooserProperty, sensorChooserBinding);

            var isListeningBinding = new Binding("IsListening") { Source = this };
            BindingOperations.SetBinding(viewModel, KinectSensorChooserUIViewModel.IsListeningProperty, isListeningBinding);

            this.expandedPopup.LayoutUpdated += this.ExpandedPopupOnLayoutUpdated;
        }

        /// <summary>
        /// Set this to true when the application is listining for audio input from the sensor.
        /// UI will show a microphone icon.  Value is passed through to the view model.
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
        /// The KinectSensorChooser whose status we are displaying.  Value is
        /// passed through to the view model.
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
        /// Used internally to transfer the visual state from the view model to the
        /// VisualStateManager defined in XAML.
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

        private void OnLoaded(object sender, RoutedEventArgs routedEventArgs)
        {
            this.parentWindow = Window.GetWindow(this);
            if (parentWindow != null)
            {
                parentWindow.Deactivated += ParentWindowOnDeactivated;
            }
        }

        private void OnUnloaded(object sender, RoutedEventArgs routedEventArgs)
        {
            if (this.parentWindow != null)
            {
                parentWindow.Deactivated -= ParentWindowOnDeactivated;
            }
        }

        private void ParentWindowOnDeactivated(object sender, EventArgs eventArgs)
        {
            this.ClosePopup();
        }

        private void ClosePopup()
        {
            this.expandedPopup.IsOpen = false;
        }

        private void OpenPopup()
        {
            this.expandedPopup.IsOpen = true;
        }

        private void OnRootGridGotKeyboardFocus(object sender, RoutedEventArgs e)
        {
            if (!this.suppressPopupOnFocus)
            {
                OpenPopup();
            }
        }

        private void OnRootGridMouseEnter(object sender, MouseEventArgs e)
        {
            OpenPopup();
        }

        private void ExpandedPopupOnLayoutUpdated(object sender, EventArgs eventArgs)
        {
            // makes the popup top-aligned with its parent
            this.expandedPopup.VerticalOffset = (this.popupGrid.ActualHeight - this.layoutRoot.ActualHeight - 1.0) / 2.0;
        }

        private void ExpandedPopupOnOpened(object sender, EventArgs eventArgs)
        {
            this.popupCloseCheck.Stop();

            if (this.layoutRoot.IsKeyboardFocusWithin)
            {
                Keyboard.Focus(this.popupGrid);
            }
            else
            {
                this.popupCloseCheck.Start();
            }
        }

        private void OnExpandedPopupMouseLeave(object sender, MouseEventArgs e)
        {
            this.ClosePopup();
        }

        private void OnPopupCloseCheckFired(object sender, EventArgs e)
        {
            this.popupCloseCheck.Stop();
            if (this.expandedPopup.IsOpen && !this.popupGrid.IsMouseOver && !this.popupGrid.IsKeyboardFocusWithin)
            {
                this.ClosePopup();
            }
        }

        private void OnPopupGridGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            var oldFocus = e.OldFocus as FrameworkElement;
            var newFocus = e.NewFocus as FrameworkElement;

            if (newFocus == this.popupGrid)
            {
                if (oldFocus != this.layoutRoot)
                {
                    // Focus is returning to us after being tabbed around our children.
                    // That is our signal to quit the popup.
                    this.suppressPopupOnFocus = true;
                    this.layoutRoot.Focusable = false;
                    e.Handled = true;

                    // There doesn't seem to be an easy way to generically tell which way the user was
                    // navigating with the keyboard (tab or shift+tab) check if the shift key is pressed
                    FocusNavigationDirection direction = ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
                                                             ? FocusNavigationDirection.Previous
                                                             : FocusNavigationDirection.Next;

                    this.ClosePopup();
                    this.MoveFocus(new TraversalRequest(direction));
                    this.layoutRoot.Focusable = true;
                    this.suppressPopupOnFocus = false;
                }
            }
        }

        private void OnVisualstateChanged(string newState)
        {
            VisualStateManager.GoToState(this, newState, true);
        }

        private void TellMeMoreLinkRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            var hyperlink = e.OriginalSource as Hyperlink;
            if (hyperlink != null)
            {
                try
                {
                    // Careful - ensure that this NavigateUri comes from a trusted source, as in this sample, before launching a process using it.
                    Process.Start(new ProcessStartInfo(hyperlink.NavigateUri.ToString()));
                }
                catch (Win32Exception)
                {
                    // No default browser was set to handle the http request or unable to launch the browser
                    MessageBox.Show(string.Format(CultureInfo.CurrentCulture, Properties.Resources.NoDefaultBrowserAvailable, hyperlink.NavigateUri));
                }

                this.ClosePopup();
            }

            e.Handled = true;
        }
    }
}