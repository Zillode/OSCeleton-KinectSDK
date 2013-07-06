//------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace Microsoft.Samples.Kinect.SkeletonBasics
{
    using System;
    using System.IO;
    using System.Windows.Media;
    using System.Windows.Threading;
    using System.Collections.Generic;
    using System.Diagnostics;
    using Ventuz.OSC;
    using Microsoft.Kinect;
    using Microsoft.Kinect.Toolkit.FaceTracking;
    using Microsoft.Speech.AudioFormat;
    using Microsoft.Speech.Recognition;
    using System.Linq;
    using System.Windows.Media.Animation;
    using System.Windows;
    using System.Threading;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        // Settings
        private bool allUsers = true;
        private bool fullBody = true;
        private bool faceTracking = false;
        private bool faceTracking2DMesh = false;
        private bool faceTrackingHeadPose = true;
        private bool faceTrackingAnimationUnits = true;
        private bool writeOSC = true;
        private bool writeCSV = true;
        private bool useUnixEpochTime = true;
        private String oscHost = "127.0.0.1";
        private int oscPort = 7110;
        private const int skeletonCount = 6;
        private const int pointScale = 1000;

        // Options
        private bool showSkeleton = true;
        private bool speechCommands = true;

        // Outputs
        private bool capturing = true;
        private UdpWriter osc;
        private StreamWriter fileWriter;
        private Stopwatch stopwatch;
        private static List<String> oscMapping = new List<String> { "",
            "head", "neck", "torso", "waist",
            "l_collar", "l_shoulder", "l_elbow", "l_wrist", "l_hand", "l_fingertip",
            "r_collar", "r_shoulder", "r_elbow", "r_wrist", "r_hand", "r_fingertip",
            "l_hip", "l_knee", "l_ankle", "l_foot",
            "r_hip", "r_knee", "r_ankle", "r_foot" };

        // Active values
        private Dictionary<int, Microsoft.Kinect.Toolkit.FaceTracking.FaceTracker> faceTrackers = new Dictionary<int, Microsoft.Kinect.Toolkit.FaceTracking.FaceTracker>();
        private Skeleton[] skeletons = new Skeleton[skeletonCount];
        private Byte[] colorPixelData;
        private short[] depthPixelData;
        private Boolean shuttingDown;

        /// <summary>
        /// Width of output drawing
        /// </summary>
        private const float RenderWidth = 640.0f;

        /// <summary>
        /// Height of our output drawing
        /// </summary>
        private const float RenderHeight = 480.0f;

        /// <summary>
        /// Thickness of drawn joint lines
        /// </summary>
        private const double JointThickness = 3;

        /// <summary>
        /// Thickness of body center ellipse
        /// </summary>
        private const double BodyCenterThickness = 10;

        /// <summary>
        /// Thickness of clip edge rectangles
        /// </summary>
        private const double ClipBoundsThickness = 10;

        /// <summary>
        /// Brush used to draw skeleton center point
        /// </summary>
        private readonly Brush centerPointBrush = Brushes.Blue;

        /// <summary>
        /// Brush used for drawing joints that are currently tracked
        /// </summary>
        private readonly Brush trackedJointBrush = new SolidColorBrush(Color.FromArgb(255, 68, 192, 68));

        /// <summary>
        /// Brush used for drawing joints that are currently inferred
        /// </summary>        
        private readonly Brush inferredJointBrush = Brushes.Yellow;

        /// <summary>
        /// Pen used for drawing bones that are currently tracked
        /// </summary>
        private readonly Pen trackedBonePen = new Pen(Brushes.Green, 6);

        /// <summary>
        /// Pen used for drawing bones that are currently inferred
        /// </summary>        
        private readonly Pen inferredBonePen = new Pen(Brushes.Gray, 1);

        /// <summary>
        /// Active Kinect sensor
        /// </summary>
        private KinectSensor sensor;

        /// <summary>
        /// Speech recognition engine using audio data from Kinect.
        /// </summary>
        private SpeechRecognitionEngine speechEngine;

        /// <summary>
        /// Drawing group for skeleton rendering output
        /// </summary>
        private DrawingGroup drawingGroup;

        /// <summary>
        /// Drawing image that we will display
        /// </summary>
        private DrawingImage imageSource;

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Draws indicators to show which edges are clipping skeleton data
        /// </summary>
        /// <param name="skeleton">skeleton to draw clipping information for</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        private static void RenderClippedEdges(Skeleton skeleton, DrawingContext drawingContext)
        {
            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Bottom))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new System.Windows.Rect(0, RenderHeight - ClipBoundsThickness, RenderWidth, ClipBoundsThickness));
            }

            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Top))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new System.Windows.Rect(0, 0, RenderWidth, ClipBoundsThickness));
            }

            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Left))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new System.Windows.Rect(0, 0, ClipBoundsThickness, RenderHeight));
            }

            if (skeleton.ClippedEdges.HasFlag(FrameEdges.Right))
            {
                drawingContext.DrawRectangle(
                    Brushes.Red,
                    null,
                    new System.Windows.Rect(RenderWidth - ClipBoundsThickness, 0, ClipBoundsThickness, RenderHeight));
            }
        }

        bool StringToBool(String msg)
        {
            msg = msg.ToLower();
            return msg.Equals("1") || msg.ToLower().Equals("true");
        }

        /// <summary>
        /// Execute startup tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowLoaded(object sender, System.Windows.RoutedEventArgs e)
        {
            // Install Shortcut
            CheckForShortcut();

            // Parse commandline arguments
            string[] args = Environment.GetCommandLineArgs();
            for (int index = 1; index < args.Length; index += 2)
            {
                args[index] = args[index].ToLower();
                if ("allUsers".ToLower().Equals(args[index])) allUsers = StringToBool(args[index+1]);
                if ("fullBody".ToLower().Equals(args[index])) fullBody = StringToBool(args[index + 1]);
                if ("faceTracking".ToLower().Equals(args[index])) faceTracking = StringToBool(args[index + 1]);
                if ("faceTracking2DMesh".ToLower().Equals(args[index])) faceTracking2DMesh = StringToBool(args[index + 1]);
                if ("faceTrackingHeadPose".ToLower().Equals(args[index])) faceTrackingHeadPose = StringToBool(args[index + 1]);
                if ("faceTrackingAnimationUnits".ToLower().Equals(args[index])) faceTrackingAnimationUnits = StringToBool(args[index + 1]);
                if ("writeOSC".ToLower().Equals(args[index])) writeOSC = StringToBool(args[index + 1]);
                if ("writeCSV".ToLower().Equals(args[index])) writeCSV = StringToBool(args[index + 1]);
                if ("useUnixEpochTime".ToLower().Equals(args[index])) useUnixEpochTime = StringToBool(args[index + 1]);
                if ("oscHost".ToLower().Equals(args[index])) oscHost = args[index + 1];
                if ("oscPort".ToLower().Equals(args[index]))
                {
                    if (!int.TryParse(args[index+1], out oscPort)) {
                        System.Windows.MessageBox.Show("Failed to parse the oscPort argument: " + args[index + 1]);
                    }
                }
                if ("showSkeleton".ToLower().Equals(args[index])) showSkeleton = StringToBool(args[index + 1]);
            }
            
            // Initialisation
            shuttingDown = false;
            stopwatch = new Stopwatch();
            stopwatch.Reset();
            stopwatch.Start();
            if (writeOSC)
            {
                osc = new UdpWriter(oscHost, oscPort);
            }
            if (writeCSV)
            {
                OpenNewCSVFile();
            }

            
            // Create the drawing group we'll use for drawing
            this.drawingGroup = new DrawingGroup();

            // Create an image source that we can use in our image control
            this.imageSource = new DrawingImage(this.drawingGroup);

            // Display the drawing using our image control
            Image.Source = this.imageSource;

            this.checkBoxSeatedMode.IsEnabled = false;
            this.checkBoxShowSkeleton.IsChecked = showSkeleton;
            this.checkBoxSpeechCommands.IsChecked = speechCommands;

            foreach (var potentialSensor in KinectSensor.KinectSensors)
            {
                if (potentialSensor.Status == KinectStatus.Connected)
                {
                    StartKinect(potentialSensor);
                    break;
                }
            }
             
            if (null == this.sensor)
            {
                this.statusBarText.Text = Properties.Resources.NoKinectReady;
            }

            KinectSensor.KinectSensors.StatusChanged += KinectSensorsStatusChanged; 
        }

        void KinectSensorsStatusChanged(object sender, StatusChangedEventArgs e)
        {
            switch (e.Status)
            {
                case KinectStatus.Disconnected:
                    StopKinect();
                    break;
                case KinectStatus.Connected:
                    StartKinect(e.Sensor);
                    break;
                case KinectStatus.NotReady:
                case KinectStatus.Initializing:
                    break;
                default:
                    System.Windows.MessageBox.Show("Kinect warning: " + e.Status);
                    break;
            }
        }

        void StartKinect(KinectSensor sensor)
        {
            this.sensor = sensor;
            if (this.sensor == null)
            {
                return;
            }

            var parameters = new TransformSmoothParameters
            {
                // as the smoothing value is increased responsiveness to the raw data
                // decreases; therefore, increased smoothing leads to increased latency.
                Smoothing = 0.1f,
                // higher value corrects toward the raw data more quickly,
                // a lower value corrects more slowly and appears smoother.
                Correction = 0.1f,
                // number of frames to predict into the future.
                Prediction = 0.1f,
                // determines how aggressively to remove jitter from the raw data.
                JitterRadius = 0.01f,
                // maximum radius (in meters) that filtered positions can deviate from raw data.
                MaxDeviationRadius = 0.04f
            };

            if (faceTracking)
            {
                this.sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                this.sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
                this.sensor.DepthStream.Range = DepthRange.Default;
                this.sensor.SkeletonStream.EnableTrackingInNearRange = false;
                this.sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                this.sensor.SkeletonStream.Enable(parameters);
                colorPixelData = new byte[this.sensor.ColorStream.FramePixelDataLength];
                depthPixelData = new short[this.sensor.DepthStream.FramePixelDataLength];
                sensor.AllFramesReady += SensorAllFramesReady;
            }
            else
            {
                sensor.ColorStream.Disable();
                sensor.DepthStream.Disable();
                sensor.SkeletonStream.Enable();
                sensor.DepthStream.Range = DepthRange.Default;
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                sensor.SkeletonStream.Enable(parameters);
                sensor.SkeletonFrameReady += SensorSkeletonFrameReady;
            }

            try
            {
                sensor.Start();
            }
            catch (System.IO.IOException)
            {
                System.Windows.MessageBox.Show("Failed to start the Kinect Sensor");
                return;
            }

            if (checkBoxSpeechCommands.IsChecked.Value)
            {
                CreateSpeechRecognizer();
            }
            SetStatusbarText("Kinect started");
        }

        private void StopKinect()
        {
            if (this.sensor != null)
            {
                if (this.sensor.IsRunning)
                {
                    SetStatusbarText("Kinect stopped", Colors.Red);
                    CloseCSVFile();
                    this.sensor.ColorStream.Disable();
                    this.sensor.DepthStream.Disable();
                    this.sensor.SkeletonStream.Disable();
                    StopSpeechRecognizer();
                    if (this.sensor.AudioSource != null)
                    {
                        this.sensor.AudioSource.Stop();
                    }
                    this.sensor.Stop();
                }
            }
        }


        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            shuttingDown = true;
            StopKinect();
        }

        /// <summary>
        /// Event handler for Kinect sensor's SkeletonFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorSkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            if (shuttingDown)
            {
                return;
            }
            SkeletonFrame skeletonFrame = e.OpenSkeletonFrame();
            if (skeletonFrame == null) return;
            skeletonFrame.CopySkeletonDataTo(skeletons);
            skeletonFrame = null;
            SensorFrameHelper(false);
        }
            
        /// <summary>
        /// Event handler for Kinect sensor's SkeletonFrameReady event
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void SensorAllFramesReady(object sender, AllFramesReadyEventArgs e)
        {
            if (shuttingDown)
            {
                return;
            }
            ColorImageFrame colorFrame = e.OpenColorImageFrame();
            if (colorFrame == null) return;
            colorFrame.CopyPixelDataTo(colorPixelData);
            colorFrame = null;
            DepthImageFrame depthFrame = e.OpenDepthImageFrame();
            if (depthFrame == null) return;
            depthFrame.CopyPixelDataTo(depthPixelData);
            depthFrame = null;
            SkeletonFrame skeletonFrame = e.OpenSkeletonFrame();
            if (skeletonFrame == null) return;
            skeletonFrame.CopySkeletonDataTo(skeletons);
            skeletonFrame = null;
            SensorFrameHelper(true);
        }
        
        private void SensorFrameHelper(Boolean allFrames) {
            if (capturing)
            {
                foreach (Skeleton skel in skeletons)
                {
                    if (skel.TrackingState == SkeletonTrackingState.Tracked)
                    {
                        SendSkeleton(skel.TrackingId, skel);
                        if (allFrames && faceTracking)
                            SendFaceTracking(skel.TrackingId, skel);
                    }
                }
            }

            if (showSkeleton)
            {
                using (DrawingContext dc = this.drawingGroup.Open())
                {
                    // Draw a transparent background to set the render size
                    dc.DrawRectangle(Brushes.Black, null, new System.Windows.Rect(0.0, 0.0, RenderWidth, RenderHeight));

                    if (skeletons.Length != 0)
                    {
                        foreach (Skeleton skel in skeletons)
                        {
                            RenderClippedEdges(skel, dc);

                            if (skel.TrackingState == SkeletonTrackingState.Tracked)
                            {
                                this.DrawBonesAndJoints(skel, dc);
                            }
                            else if (skel.TrackingState == SkeletonTrackingState.PositionOnly)
                            {
                                dc.DrawEllipse(
                                    this.centerPointBrush,
                                    null,
                                    this.SkeletonPointToScreen(skel.Position),
                                BodyCenterThickness,
                                BodyCenterThickness);
                            }
                        }
                    }

                    // prevent drawing outside of our render area
                    this.drawingGroup.ClipGeometry = new RectangleGeometry(new System.Windows.Rect(0.0, 0.0, RenderWidth, RenderHeight));
                }
            }
        }

        private void CreateSpeechRecognizer()
        {
            RecognizerInfo ri = SpeechRecognitionEngine.InstalledRecognizers()
                .Where(r => r.Culture.Name == "en-US").FirstOrDefault();
            if (ri == null)
            {
                return;
            }
            this.speechEngine = new SpeechRecognitionEngine(ri.Id);
            var words = new Choices();
            words.Add("start");
            words.Add("next");
            words.Add("continue");
            words.Add("pause");
            words.Add("stop");
            var gb = new GrammarBuilder();
            gb.Culture = ri.Culture;
            gb.Append(words);
            var g = new Grammar(gb);
            this.speechEngine.LoadGrammar(g);
            this.speechEngine.SpeechRecognized += SpeechRecognized;
            // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
            // This will prevent recognition accuracy from degrading over time.
            this.speechEngine.UpdateRecognizerSetting("AdaptationOn", 0);
            this.speechEngine.SetInputToAudioStream(this.sensor.AudioSource.Start(), new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            this.speechEngine.RecognizeAsync(RecognizeMode.Multiple);
            this.helpbox.Text = "Keyboard shortcuts: space; Speech commands: start, next, pause, continue, stop.";
        }

        private void StopSpeechRecognizer()
        {
            if (this.speechEngine != null)
            {
                ThreadPool.QueueUserWorkItem((object x) => (x as IDisposable).Dispose(), this.speechEngine);
                this.speechEngine = null;
            }
        }

        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Confidence < 0.05)
            {
                return;
            }

            switch (e.Result.Text.ToUpper())
            {
                case "START":
                    foreach (var potentialSensor in KinectSensor.KinectSensors)
                    {
                        if (potentialSensor.Status == KinectStatus.Connected)
                        {
                            StartKinect(potentialSensor);
                            break;
                        }
                    }
                    break;
                case "NEXT":
                    if (writeCSV)
                    {
                        capturing = true;
                        OpenNewCSVFile();
                    }
                    break;
                case "CONTINUE":
                    capturing = true;
                    SetStatusbarText("Continue capturing", Colors.Green);
                    break;
                case "PAUSE":
                    capturing = false;
                    SetStatusbarText("Paused capturing", Colors.Red);
                    break;
                case "STOP":
                    SetStatusbarText("Kinect stopped", Colors.Red);
                    StopKinect();
                    break;
            }
        }

        private void OpenNewCSVFile()
        {
            CloseCSVFile();
            fileWriter = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/points-MSK-" + getUnixEpochTime().ToString().Replace(",", ".") + ".csv", false);
            fileWriter.WriteLine("Joint, user, joint, x, y, z, confidence, time");
            fileWriter.WriteLine("Face, user, x, y, z, pitch, yaw, roll, time");
            fileWriter.WriteLine("FaceAnimation, face, lip_raise, lip_stretcher, lip_corner_depressor, jaw_lower, brow_lower, brow_raise, time");
            SetStatusbarText( "Writing to file " + DateTime.Now.ToLongTimeString(), Colors.Orange);
        }

        private void CloseCSVFile()
        {
            if (fileWriter != null)
            {
                fileWriter.Close();
                fileWriter = null;
            }
        }

        private void SetStatusbarText(string txt)
        {
            statusBarText.Text = txt;
        }

        private void SetStatusbarText(string txt, Color color)
        {
            ColorAnimation animation;
            animation = new ColorAnimation();
            animation.From = color;
            animation.To = Colors.White;
            animation.Duration = new Duration(TimeSpan.FromSeconds(3));
            statusBar.Background = new SolidColorBrush(color);
            statusBar.Background.BeginAnimation(SolidColorBrush.ColorProperty, animation);
            statusBarText.Text = txt;
        }

        /// <summary>
        /// Draws a skeleton's bones and joints
        /// </summary>
        /// <param name="skeleton">skeleton to draw</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        private void DrawBonesAndJoints(Skeleton skeleton, DrawingContext drawingContext)
        {
            // Render Torso
            this.DrawBone(skeleton, drawingContext, JointType.Head, JointType.ShoulderCenter);
            this.DrawBone(skeleton, drawingContext, JointType.ShoulderCenter, JointType.ShoulderLeft);
            this.DrawBone(skeleton, drawingContext, JointType.ShoulderCenter, JointType.ShoulderRight);
            this.DrawBone(skeleton, drawingContext, JointType.ShoulderCenter, JointType.Spine);
            this.DrawBone(skeleton, drawingContext, JointType.Spine, JointType.HipCenter);
            this.DrawBone(skeleton, drawingContext, JointType.HipCenter, JointType.HipLeft);
            this.DrawBone(skeleton, drawingContext, JointType.HipCenter, JointType.HipRight);

            // Left Arm
            this.DrawBone(skeleton, drawingContext, JointType.ShoulderLeft, JointType.ElbowLeft);
            this.DrawBone(skeleton, drawingContext, JointType.ElbowLeft, JointType.WristLeft);
            this.DrawBone(skeleton, drawingContext, JointType.WristLeft, JointType.HandLeft);

            // Right Arm
            this.DrawBone(skeleton, drawingContext, JointType.ShoulderRight, JointType.ElbowRight);
            this.DrawBone(skeleton, drawingContext, JointType.ElbowRight, JointType.WristRight);
            this.DrawBone(skeleton, drawingContext, JointType.WristRight, JointType.HandRight);

            // Left Leg
            this.DrawBone(skeleton, drawingContext, JointType.HipLeft, JointType.KneeLeft);
            this.DrawBone(skeleton, drawingContext, JointType.KneeLeft, JointType.AnkleLeft);
            this.DrawBone(skeleton, drawingContext, JointType.AnkleLeft, JointType.FootLeft);

            // Right Leg
            this.DrawBone(skeleton, drawingContext, JointType.HipRight, JointType.KneeRight);
            this.DrawBone(skeleton, drawingContext, JointType.KneeRight, JointType.AnkleRight);
            this.DrawBone(skeleton, drawingContext, JointType.AnkleRight, JointType.FootRight);
 
            // Render Joints
            foreach (Joint joint in skeleton.Joints)
            {
                Brush drawBrush = null;

                if (joint.TrackingState == JointTrackingState.Tracked)
                {
                    drawBrush = this.trackedJointBrush;                    
                }
                else if (joint.TrackingState == JointTrackingState.Inferred)
                {
                    drawBrush = this.inferredJointBrush;                    
                }

                if (drawBrush != null)
                {
                    drawingContext.DrawEllipse(drawBrush, null, this.SkeletonPointToScreen(joint.Position), JointThickness, JointThickness);
                }
            }
        }

        /// <summary>
        /// Maps a SkeletonPoint to lie within our render space and converts to Point
        /// </summary>
        /// <param name="skelpoint">point to map</param>
        /// <returns>mapped point</returns>
        private System.Windows.Point SkeletonPointToScreen(SkeletonPoint skelpoint)
        {
            // Convert point to depth space.  
            // We are not using depth directly, but we do want the points in our 640x480 output resolution.
            DepthImagePoint depthPoint = this.sensor.CoordinateMapper.MapSkeletonPointToDepthPoint(skelpoint, DepthImageFormat.Resolution640x480Fps30);
            return new System.Windows.Point(depthPoint.X, depthPoint.Y);
        }

        /// <summary>
        /// Draws a bone line between two joints
        /// </summary>
        /// <param name="skeleton">skeleton to draw bones from</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        /// <param name="jointType0">joint to start drawing from</param>
        /// <param name="jointType1">joint to end drawing at</param>
        private void DrawBone(Skeleton skeleton, DrawingContext drawingContext, JointType jointType0, JointType jointType1)
        {
            Joint joint0 = skeleton.Joints[jointType0];
            Joint joint1 = skeleton.Joints[jointType1];

            // If we can't find either of these joints, exit
            if (joint0.TrackingState == JointTrackingState.NotTracked ||
                joint1.TrackingState == JointTrackingState.NotTracked)
            {
                return;
            }

            // Don't draw if both points are inferred
            if (joint0.TrackingState == JointTrackingState.Inferred &&
                joint1.TrackingState == JointTrackingState.Inferred)
            {
                return;
            }

            // We assume all drawn bones are inferred unless BOTH joints are tracked
            Pen drawPen = this.inferredBonePen;
            if (joint0.TrackingState == JointTrackingState.Tracked && joint1.TrackingState == JointTrackingState.Tracked)
            {
                drawPen = this.trackedBonePen;
            }

            drawingContext.DrawLine(drawPen, this.SkeletonPointToScreen(joint0.Position), this.SkeletonPointToScreen(joint1.Position));
        }

        /// <summary>
        /// Handles the checking or unchecking of the seated mode combo box
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void CheckBoxSeatedModeChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            if (null != this.sensor)
            {
                if (this.checkBoxSeatedMode.IsChecked.GetValueOrDefault())
                {
                    this.sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                }
                else
                {
                    this.sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                }
            }
        }

        /// <summary>
        /// Handles the checking or unchecking of the seated mode combo box
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void CheckBoxShowSkeletonChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            this.showSkeleton = this.checkBoxShowSkeleton.IsChecked.GetValueOrDefault();
        }


        void SendFaceTracking(int user, Skeleton s)
        {
            if (!faceTrackers.ContainsKey(user))
            {
                if (faceTrackers.Count > 10)
                {
                    faceTrackers.Clear();
                }
                try
                {
                    faceTrackers.Add(user, new Microsoft.Kinect.Toolkit.FaceTracking.FaceTracker(sensor));
                }
                catch (InvalidOperationException)
                {
                    // During some shutdown scenarios the FaceTracker
                    // is unable to be instantiated.  Catch that exception
                    // and don't track a face.
                    Debug.WriteLine("AllFramesReady - creating a new FaceTracker threw an InvalidOperationException");
                    return;
                }
            }

            Microsoft.Kinect.Toolkit.FaceTracking.FaceTrackFrame faceFrame = faceTrackers[user].Track(
                            sensor.ColorStream.Format, colorPixelData,
                            sensor.DepthStream.Format, depthPixelData,
                            s);

            if (faceFrame.TrackSuccessful)
            {
                if (faceTracking2DMesh)
                {
                    // TODO
                    // faceFrame.Get3DShape[FeaturePoint.]
                }

                if (faceTrackingHeadPose)
                {
                    SendFacePoseMessage(user,
                        s.Joints[JointType.Head].Position.X, s.Joints[JointType.Head].Position.Y, s.Joints[JointType.Head].Position.Z,
                        faceFrame.Rotation.X, faceFrame.Rotation.Y, faceFrame.Rotation.Z, useUnixEpochTime ? getUnixEpochTime() : stopwatch.ElapsedMilliseconds);
                }

                if (faceTrackingAnimationUnits)
                {
                    SendFaceAnimationMessage(user, faceFrame.GetAnimationUnitCoefficients(), useUnixEpochTime ? getUnixEpochTime() : stopwatch.ElapsedMilliseconds);
                }
            }
        }

        void SendSkeleton(int user, Skeleton s)
        {
            if (!capturing) { return; }

            if (!fullBody)
            {
                ProcessJointInformation(user, 15, s.Joints[JointType.HandRight], s.BoneOrientations[JointType.HandRight]);
            }
            else
            {
                ProcessJointInformation(user, 1, s.Joints[JointType.Head], s.BoneOrientations[JointType.Head]);
                ProcessJointInformation(user, 2, s.Joints[JointType.ShoulderCenter], s.BoneOrientations[JointType.ShoulderCenter]);
                ProcessJointInformation(user, 3, s.Joints[JointType.Spine], s.BoneOrientations[JointType.Spine]);
                ProcessJointInformation(user, 4, s.Joints[JointType.HipCenter], s.BoneOrientations[JointType.HipCenter]);
                // ProcessJointInformation(user, 5, s.Joints[JointType.], s.BoneOrientations[JointType.]);
                ProcessJointInformation(user, 6, s.Joints[JointType.ShoulderLeft], s.BoneOrientations[JointType.ShoulderLeft]);
                ProcessJointInformation(user, 7, s.Joints[JointType.ElbowLeft], s.BoneOrientations[JointType.ElbowLeft]);
                ProcessJointInformation(user, 8, s.Joints[JointType.WristLeft], s.BoneOrientations[JointType.WristLeft]);
                ProcessJointInformation(user, 9, s.Joints[JointType.HandLeft], s.BoneOrientations[JointType.HandLeft]);
                // ProcessJointInformation(user, 10, s.Joints[JointType.], s.BoneOrientations[JointType.]);
                // ProcessJointInformation(user, 11, s.Joints[JointType.], s.BoneOrientations[JointType.]);
                ProcessJointInformation(user, 12, s.Joints[JointType.ShoulderRight], s.BoneOrientations[JointType.ShoulderRight]);
                ProcessJointInformation(user, 13, s.Joints[JointType.ElbowRight], s.BoneOrientations[JointType.ElbowRight]);
                ProcessJointInformation(user, 14, s.Joints[JointType.WristRight], s.BoneOrientations[JointType.WristRight]);
                ProcessJointInformation(user, 15, s.Joints[JointType.HandRight], s.BoneOrientations[JointType.HandRight]);
                // ProcessJointInformation(user, 16, s.Joints[JointType.], s.BoneOrientations[JointType.]);
                ProcessJointInformation(user, 17, s.Joints[JointType.HipLeft], s.BoneOrientations[JointType.HipLeft]);
                ProcessJointInformation(user, 18, s.Joints[JointType.KneeLeft], s.BoneOrientations[JointType.KneeLeft]);
                ProcessJointInformation(user, 19, s.Joints[JointType.AnkleLeft], s.BoneOrientations[JointType.AnkleLeft]);
                ProcessJointInformation(user, 20, s.Joints[JointType.FootLeft], s.BoneOrientations[JointType.FootLeft]);
                ProcessJointInformation(user, 21, s.Joints[JointType.HipRight], s.BoneOrientations[JointType.HipRight]);
                ProcessJointInformation(user, 22, s.Joints[JointType.KneeRight], s.BoneOrientations[JointType.KneeRight]);
                ProcessJointInformation(user, 23, s.Joints[JointType.AnkleRight], s.BoneOrientations[JointType.AnkleRight]);
                ProcessJointInformation(user, 24, s.Joints[JointType.FootRight], s.BoneOrientations[JointType.FootRight]);
            }
        }

        double JointToConfidenceValue(Joint j)
        {
            if (j.TrackingState == JointTrackingState.Tracked) return 1;
            if (j.TrackingState == JointTrackingState.Inferred) return 0.5;
            if (j.TrackingState == JointTrackingState.NotTracked) return 0.1;
            return 0.5;
        }

        void ProcessJointInformation(int user, int joint, Joint j, BoneOrientation bo)
        {
            SendJointMessage(user, joint,
                j.Position.X, j.Position.Y, j.Position.Z,
                JointToConfidenceValue(j),
                stopwatch.ElapsedMilliseconds);
        }

        void SendJointMessage(int user, int joint, double x, double y, double z, double confidence, double time)
        {
            if (osc != null)
            {
                osc.Send(new OscElement("/joint", oscMapping[joint], user, (float)(x * pointScale), (float)(-y * pointScale), (float)(z * pointScale), (float)confidence, time));
            }
            if (fileWriter != null)
            {
                // Joint, user, joint, x, y, z, on
                fileWriter.WriteLine("Joint," + user + "," + joint + "," +
                    (x * pointScale).ToString().Replace(",", ".") + "," +
                    (-y * pointScale).ToString().Replace(",", ".") + "," +
                    (z * pointScale).ToString().Replace(",", ".") + "," +
                    confidence.ToString().Replace(",", ".") + "," +
                    time.ToString().Replace(",", "."));
            }
        }

        void SendFacePoseMessage(int user, float x, float y, float z, float rotationX, float rotationY, float rotationZ, double time)
        {
            if (osc != null)
            {
                osc.Send(new OscElement(
                    "/face",
                    user,
                    (float)(x * pointScale), (float)(-y * pointScale), (float)(z * pointScale),
                    rotationX, rotationY, rotationZ,
                    time));
            }
            if (fileWriter != null)
            {
                fileWriter.WriteLine("Face," +
                    user + "," +
                    (x * pointScale).ToString().Replace(",", ".") + "," +
                    (-y * pointScale).ToString().Replace(",", ".") + "," +
                    (z * pointScale).ToString().Replace(",", ".") + "," +
                    rotationX.ToString().Replace(",", ".") + "," +
                    rotationY.ToString().Replace(",", ".") + "," +
                    rotationZ.ToString().Replace(",", ".") + "," +
                    time.ToString().Replace(",", "."));
            }
        }

        void SendFaceAnimationMessage(int user, EnumIndexableCollection<AnimationUnit, float> c, double time)
        {
            if (osc != null)
            {
                osc.Send(new OscElement(
                    "/face_animation",
                    user,
                    c[AnimationUnit.LipRaiser],
                    c[AnimationUnit.LipStretcher],
                    c[AnimationUnit.LipCornerDepressor],
                    c[AnimationUnit.JawLower],
                    c[AnimationUnit.BrowLower],
                    c[AnimationUnit.BrowRaiser],
                    time));
            }
            if (fileWriter != null)
            {
                fileWriter.WriteLine("FaceAnimation," +
                    user + "," +
                    c[AnimationUnit.LipRaiser].ToString().Replace(",", ".") + "," +
                    c[AnimationUnit.LipStretcher].ToString().Replace(",", ".") + "," +
                    c[AnimationUnit.LipCornerDepressor].ToString().Replace(",", ".") + "," +
                    c[AnimationUnit.JawLower].ToString().Replace(",", ".") + "," +
                    c[AnimationUnit.BrowLower].ToString().Replace(",", ".") + "," +
                    c[AnimationUnit.BrowRaiser].ToString().Replace(",", ".") + "," +
                    time.ToString().Replace(",", "."));
            }
        }

        private string GetErrorText(Exception ex)
        {
            string err = ex.Message;
            if (ex.InnerException != null)
            {
                err += " - More details: " + ex.InnerException.Message;
            }
            return err;
        }

        public void CheckForShortcut()
        {
            try
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    return;
                }
                System.Deployment.Application.ApplicationDeployment ad = default(System.Deployment.Application.ApplicationDeployment);
                ad = System.Deployment.Application.ApplicationDeployment.CurrentDeployment;

                if ((ad.IsFirstRun))
                {
                    System.Reflection.Assembly code = System.Reflection.Assembly.GetExecutingAssembly();
                    string company = string.Empty;
                    string description = string.Empty;

                    if ((Attribute.IsDefined(code, typeof(System.Reflection.AssemblyCompanyAttribute))))
                    {
                        System.Reflection.AssemblyCompanyAttribute ascompany = null;
                        ascompany = (System.Reflection.AssemblyCompanyAttribute)Attribute.GetCustomAttribute(code, typeof(System.Reflection.AssemblyCompanyAttribute));
                        company = ascompany.Company;
                    }

                    if ((Attribute.IsDefined(code, typeof(System.Reflection.AssemblyTitleAttribute))))
                    {
                        System.Reflection.AssemblyTitleAttribute asdescription = null;
                        asdescription = (System.Reflection.AssemblyTitleAttribute)Attribute.GetCustomAttribute(code, typeof(System.Reflection.AssemblyTitleAttribute));
                        description = asdescription.Title;

                    }

                    if ((company != string.Empty & description != string.Empty))
                    {
                        //description = Replace(description, "_", " ")

                        string desktopPath = string.Empty;
                        desktopPath = string.Concat(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "\\", description, ".appref-ms");

                        string shortcutName = string.Empty;
                        shortcutName = string.Concat(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "\\", company, "\\", description, ".appref-ms");

                        System.IO.File.Copy(shortcutName, desktopPath, true);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("Missing company or description: " + company + " - " + description);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(GetErrorText(ex));
            }
        }

        private double getUnixEpochTime()
        {
            var unixTime = DateTime.Now.ToUniversalTime() - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return unixTime.TotalMilliseconds;
        }

        private long getUnixEpochTimeLong()
        {
            var unixTime = DateTime.Now.ToUniversalTime() - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return System.Convert.ToInt64(unixTime.TotalMilliseconds);
        }

        private void Window_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Space)
            {
                if (writeCSV)
                {
                    OpenNewCSVFile();
                }
            }
        }

        private void CheckBoxSpeechCommandsChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            if (this.sensor == null) return;
            if (this.checkBoxSpeechCommands.IsChecked.Value)
            {
                CreateSpeechRecognizer();
            } else {
                StopSpeechRecognizer();
            }
        }

    }
}