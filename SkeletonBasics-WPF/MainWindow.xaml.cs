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
    using System.Windows.Media.Imaging;
    using System.Collections.Concurrent;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        public class SkeletonInformation
        {
            public int sensorId;
            public int user;
            public Skeleton skeleton;
            public long time;

            public SkeletonInformation(int sensorId, int user, Skeleton skeleton, long time)
            {
                this.sensorId = sensorId;
                this.user = user;
                this.skeleton = skeleton;
                this.time = time;
            }
        }

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
        private bool showRGB = false;
        private bool showDepth = false;
        private bool speechCommands = true;

        // Outputs
        private int sensorId = 0;
        private bool capturing = true;
        private BlockingCollection<SkeletonInformation> skeletonInformationQueue = new BlockingCollection<SkeletonInformation>();
        Thread sendSkeleton;
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
        private List<KinectSensor> sensors = new List<KinectSensor>();
        private Dictionary<KinectSensor, int> sensorIds = new Dictionary<KinectSensor, int>();
        private Dictionary<KinectSensor, Dictionary<int, Microsoft.Kinect.Toolkit.FaceTracking.FaceTracker>> faceTrackers = new Dictionary<KinectSensor, Dictionary<int, Microsoft.Kinect.Toolkit.FaceTracking.FaceTracker>>();
        private Dictionary<KinectSensor, Skeleton[]> skeletons = new Dictionary<KinectSensor, Skeleton[]>();
        private Dictionary<KinectSensor, Byte[]> colorPixelData = new Dictionary<KinectSensor,byte[]>();
        private Dictionary<KinectSensor, short[]> depthPixelData = new Dictionary<KinectSensor,short[]>();
        private Dictionary<KinectSensor, SpeechRecognitionEngine> speechEngine = new Dictionary<KinectSensor, SpeechRecognitionEngine>();

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
        /// Drawing group for skeleton rendering output
        /// </summary>
        private DrawingGroup drawingGroup;

        /// <summary>
        /// Drawing image that we will display
        /// </summary>
        private DrawingImage imageSource;

        /// <summary>
        /// Drawing image that we will display
        /// </summary>
        private WriteableBitmap cameraSource;

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
                }
            }
             
            if (this.sensors.Count == 0)
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
                    StopKinect(e.Sensor);
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
            if (!this.sensors.Contains(sensor))
            {
                this.sensors.Add(sensor);
                this.sensorIds.Add(sensor, sensorId++);
                if (sensorId > int.MaxValue / (skeletonCount * 2))
                    sensorId = 0;
                this.skeletons.Add(sensor, new Skeleton[skeletonCount]);
                this.faceTrackers.Add(sensor, new Dictionary<int, Microsoft.Kinect.Toolkit.FaceTracking.FaceTracker>());
                this.colorPixelData.Add(sensor, null);
                this.depthPixelData.Add(sensor, null);
            }

            SetFrames(sensor);

            try
            {
                sensor.Start();
            }
            catch (System.IO.IOException)
            {
                StopKinect(sensor);
                System.Windows.MessageBox.Show("Failed to start the Kinect Sensor");
                return;
            }

            if (checkBoxSpeechCommands.IsChecked.Value)
            {
                CreateSpeechRecognizer(sensor);
            }
            if (this.sensors.Count > 1)
                SetStatusbarText("Kinect started ("+this.sensors.Count+" sensors active)");
            else
                SetStatusbarText("Kinect started");

        }

        private void SetFrames(KinectSensor sensor)
        {
            if (sensor == null) return;

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

            sensor.SkeletonFrameReady -= SensorSkeletonFrameReady;
            sensor.AllFramesReady -= SensorAllFramesReady;
            cameraSource = null;

            if (faceTracking)
            {
                if (sendSkeleton != null) {
                    sendSkeleton.Abort();
                    sendSkeleton = null;
                }
            }
            else
            {
                if (sendSkeleton == null)
                {
                    sendSkeleton = new Thread(SendSkeleton);
                    sendSkeleton.Start();
                }
            }

            if (((showRGB || showDepth) && sensors.FirstOrDefault() == sensor) || faceTracking)
            {
                sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
                sensor.DepthStream.Range = DepthRange.Default;
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                sensor.SkeletonStream.Enable(parameters);
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
        }

        private void StopKinect(KinectSensor sensor)
        {
            if (this.sensors.Contains(sensor))
            {
                this.sensors.Remove(sensor);
                this.faceTrackers.Remove(sensor);
                this.colorPixelData.Remove(sensor);
                this.depthPixelData.Remove(sensor);
            }
            if (this.sensors.Count > 0)
                SetStatusbarText("Kinect stopped (" + this.sensors.Count + " sensors active)", Colors.Red);
            else
                SetStatusbarText("Kinect stopped", Colors.Red);
            if (sensors.Count == 0)
            {
                CloseCSVFile();
            }
            if (sensor.IsRunning)
            {
                sensor.ColorStream.Disable();
                sensor.DepthStream.Disable();
                sensor.SkeletonStream.Disable();
                StopSpeechRecognizer(sensor);
                if (sensor.AudioSource != null)
                {
                    sensor.AudioSource.Stop();
                }
                sensor.Stop();
            }
        }


        /// <summary>
        /// Execute shutdown tasks
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (sendSkeleton != null)
            {
                sendSkeleton.Abort();
                sendSkeleton = null;
            }
            shuttingDown = true;
            List<KinectSensor> kinects = new List<KinectSensor>(this.sensors);
            foreach (KinectSensor kinect in kinects)
                StopKinect(kinect);
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
            KinectSensor sensor = (KinectSensor)sender;
            SkeletonFrame skeletonFrame = e.OpenSkeletonFrame();
            if (skeletonFrame == null) return;
            skeletonFrame.CopySkeletonDataTo(skeletons[sensor]);
            skeletonFrame = null;
            SensorFrameHelper(sensor, false);
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
            KinectSensor sensor = (KinectSensor)sender;
            ColorImageFrame colorFrame = e.OpenColorImageFrame();
            if (colorFrame == null) return;
            if (colorPixelData[sensor] == null)
                colorPixelData[sensor] = new byte[colorFrame.PixelDataLength];
            colorFrame.CopyPixelDataTo(colorPixelData[sensor]);
            colorFrame = null;
            DepthImageFrame depthFrame = e.OpenDepthImageFrame();
            if (depthFrame == null) return;
            if (depthPixelData[sensor] == null)
                depthPixelData[sensor] = new short[depthFrame.PixelDataLength];
            depthFrame.CopyPixelDataTo(depthPixelData[sensor]);
            depthFrame = null;
            SkeletonFrame skeletonFrame = e.OpenSkeletonFrame();
            if (skeletonFrame == null) return;
            skeletonFrame.CopySkeletonDataTo(skeletons[sensor]);
            skeletonFrame = null;
            SensorFrameHelper(sensor, true);
        }
        
        private void SensorFrameHelper(KinectSensor sensor, Boolean allFrames) {
            if (capturing)
            {
                for (int i = 0; i < skeletonCount; i++)
                {
                    Skeleton skel = skeletons[sensor][i];
                    if (skel == null) continue;
                    if (skel.TrackingState == SkeletonTrackingState.Tracked)
                    {
                        int userId = (sensorIds[sensor] * skeletonCount) + i;
                        EnqueueSkeleton(sensorIds[sensor], userId, skel);
                        if (allFrames && faceTracking)
                            SendFaceTracking(sensor, userId, skel);
                    }
                }
            }

            if ((showSkeleton || showRGB || showDepth) && sensors.FirstOrDefault() == sensor)
            {
                using (DrawingContext dc = this.drawingGroup.Open())
                {
                    // Draw a transparent background to set the render size
                    dc.DrawRectangle(Brushes.Black, null, new System.Windows.Rect(0.0, 0.0, RenderWidth, RenderHeight));

                    if (showRGB && this.colorPixelData[sensor] != null) {
                        if (cameraSource == null)
                            cameraSource = new WriteableBitmap(sensor.ColorStream.FrameWidth, sensor.ColorStream.FrameHeight,
                                96, 96, PixelFormats.Bgr32, null);
                        cameraSource.WritePixels(
                            new Int32Rect(0, 0, sensor.ColorStream.FrameWidth, sensor.ColorStream.FrameHeight),
                                this.colorPixelData[sensor],
                                sensor.ColorStream.FrameWidth * sensor.ColorStream.FrameBytesPerPixel,
                                0);
                        dc.DrawImage(cameraSource, new System.Windows.Rect(0.0, 0.0, sensor.ColorStream.FrameWidth, sensor.ColorStream.FrameHeight));
                    }

                    if (showDepth && this.depthPixelData[sensor] != null)
                    {
                        if (cameraSource == null)
                            cameraSource = new WriteableBitmap(sensor.DepthStream.FrameWidth, sensor.DepthStream.FrameHeight,
                                96, 96, PixelFormats.Gray16, null);
                        cameraSource.WritePixels(
                            new Int32Rect(0, 0, sensor.DepthStream.FrameWidth, sensor.DepthStream.FrameHeight),
                                this.depthPixelData[sensor],
                                sensor.DepthStream.FrameWidth * sensor.DepthStream.FrameBytesPerPixel,
                                0);
                        dc.DrawImage(cameraSource, new System.Windows.Rect(160, 120, sensor.DepthStream.FrameWidth, sensor.DepthStream.FrameHeight));
                    }
                    if (showSkeleton && skeletons[sensor].Length != 0)
                    {
                        foreach (Skeleton skel in skeletons[sensor])
                        {
                            RenderClippedEdges(skel, dc);

                            if (skel.TrackingState == SkeletonTrackingState.Tracked)
                            {
                                this.DrawBonesAndJoints(sensor, skel, dc);
                            }
                            else if (skel.TrackingState == SkeletonTrackingState.PositionOnly)
                            {
                                dc.DrawEllipse(
                                    this.centerPointBrush,
                                    null,
                                    this.SkeletonPointToScreen(sensor, skel.Position),
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

        private void CreateSpeechRecognizer(KinectSensor sensor)
        {
            if (this.speechEngine.ContainsKey(sensor)) return;
            RecognizerInfo ri = SpeechRecognitionEngine.InstalledRecognizers()
                .Where(r => r.Culture.Name == "en-US").FirstOrDefault();
            if (ri == null)
            {
                return;
            }
            SpeechRecognitionEngine speech = new SpeechRecognitionEngine(ri.Id);
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
            speech.LoadGrammar(g);
            speech.SpeechRecognized += SpeechRecognized;
            // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
            // This will prevent recognition accuracy from degrading over time.
            speech.UpdateRecognizerSetting("AdaptationOn", 0);
            speech.SetInputToAudioStream(sensor.AudioSource.Start(), new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            speech.RecognizeAsync(RecognizeMode.Multiple);
            this.speechEngine[sensor] = speech;
            this.helpbox.Text = "Keyboard shortcuts: space; Speech commands: start, next, pause, continue, stop.";
        }

        private void StopSpeechRecognizer(KinectSensor sensor)
        {
            if (!this.speechEngine.ContainsKey(sensor)) return;
            var speechEngine = this.speechEngine[sensor];
            this.speechEngine.Remove(sensor);
            if (speechEngine != null)
            {
                ThreadPool.QueueUserWorkItem((object x) => (x as IDisposable).Dispose(), speechEngine);
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
                    List<KinectSensor> kinects = new List<KinectSensor>(sensors);
                    foreach (KinectSensor kinect in kinects)
                        StopKinect(kinect);
                    break;
            }
        }

        private void OpenNewCSVFile()
        {
            CloseCSVFile();
            fileWriter = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/points-MSK-" + getUnixEpochTime().ToString().Replace(",", ".") + ".csv", false);
            fileWriter.WriteLine("Joint, sensor, user, joint, x, y, z, confidence, time");
            fileWriter.WriteLine("Face, sensor, user, x, y, z, pitch, yaw, roll, time");
            fileWriter.WriteLine("FaceAnimation, sensor, user, lip_raise, lip_stretcher, lip_corner_depressor, jaw_lower, brow_lower, brow_raise, time");
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
        private void DrawBonesAndJoints(KinectSensor sensor, Skeleton skeleton, DrawingContext drawingContext)
        {
            // Render Torso
            this.DrawBone(sensor, skeleton, drawingContext, JointType.Head, JointType.ShoulderCenter);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ShoulderCenter, JointType.ShoulderLeft);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ShoulderCenter, JointType.ShoulderRight);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ShoulderCenter, JointType.Spine);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.Spine, JointType.HipCenter);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.HipCenter, JointType.HipLeft);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.HipCenter, JointType.HipRight);

            // Left Arm
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ShoulderLeft, JointType.ElbowLeft);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ElbowLeft, JointType.WristLeft);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.WristLeft, JointType.HandLeft);

            // Right Arm
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ShoulderRight, JointType.ElbowRight);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.ElbowRight, JointType.WristRight);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.WristRight, JointType.HandRight);

            // Left Leg
            this.DrawBone(sensor, skeleton, drawingContext, JointType.HipLeft, JointType.KneeLeft);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.KneeLeft, JointType.AnkleLeft);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.AnkleLeft, JointType.FootLeft);

            // Right Leg
            this.DrawBone(sensor, skeleton, drawingContext, JointType.HipRight, JointType.KneeRight);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.KneeRight, JointType.AnkleRight);
            this.DrawBone(sensor, skeleton, drawingContext, JointType.AnkleRight, JointType.FootRight);
 
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
                    drawingContext.DrawEllipse(drawBrush, null, this.SkeletonPointToScreen(sensor, joint.Position), JointThickness, JointThickness);
                }
            }
        }

        /// <summary>
        /// Maps a SkeletonPoint to lie within our render space and converts to Point
        /// </summary>
        /// <param name="skelpoint">point to map</param>
        /// <returns>mapped point</returns>
        private System.Windows.Point SkeletonPointToScreen(KinectSensor sensor, SkeletonPoint skelpoint)
        {
            // Convert point to depth space.  
            // We are not using depth directly, but we do want the points in our 640x480 output resolution.
            DepthImagePoint depthPoint = sensor.CoordinateMapper.MapSkeletonPointToDepthPoint(skelpoint, DepthImageFormat.Resolution640x480Fps30);
            return new System.Windows.Point(depthPoint.X, depthPoint.Y);
        }

        /// <summary>
        /// Draws a bone line between two joints
        /// </summary>
        /// <param name="skeleton">skeleton to draw bones from</param>
        /// <param name="drawingContext">drawing context to draw to</param>
        /// <param name="jointType0">joint to start drawing from</param>
        /// <param name="jointType1">joint to end drawing at</param>
        private void DrawBone(KinectSensor sensor, Skeleton skeleton, DrawingContext drawingContext, JointType jointType0, JointType jointType1)
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

            drawingContext.DrawLine(drawPen, this.SkeletonPointToScreen(sensor, joint0.Position), this.SkeletonPointToScreen(sensor, joint1.Position));
        }

        /// <summary>
        /// Handles the checking or unchecking of the seated mode combo box
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void CheckBoxSeatedModeChanged(object sender, System.Windows.RoutedEventArgs e)
        {
            List<KinectSensor> kinects = new List<KinectSensor>(sensors);
            foreach (KinectSensor kinect in kinects) {
                if (this.checkBoxSeatedMode.IsChecked.GetValueOrDefault())
                {
                    kinect.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                }
                else
                {
                    kinect.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
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


        void SendFaceTracking(KinectSensor sensor, int user, Skeleton s)
        {
            Dictionary<int, FaceTracker> faceTrackers = this.faceTrackers[sensor];
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
                            sensor.ColorStream.Format, colorPixelData[sensor],
                            sensor.DepthStream.Format, depthPixelData[sensor],
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
                    SendFacePoseMessage(sensorId, user,
                        s.Joints[JointType.Head].Position.X, s.Joints[JointType.Head].Position.Y, s.Joints[JointType.Head].Position.Z,
                        faceFrame.Rotation.X, faceFrame.Rotation.Y, faceFrame.Rotation.Z, useUnixEpochTime ? getUnixEpochTime() : stopwatch.ElapsedMilliseconds);
                }

                if (faceTrackingAnimationUnits)
                {
                    SendFaceAnimationMessage(sensorId, user, faceFrame.GetAnimationUnitCoefficients(), useUnixEpochTime ? getUnixEpochTime() : stopwatch.ElapsedMilliseconds);
                }
            }
        }

        void EnqueueSkeleton(int sensorId, int user, Skeleton s)
        {
            skeletonInformationQueue.Add(new SkeletonInformation(sensorId, user, s, stopwatch.ElapsedMilliseconds));
        }
          
        void SendSkeleton() {
            while (true) {
                SkeletonInformation i = skeletonInformationQueue.Take();
                SendSkeleton2(i.sensorId, i.user, i.skeleton, i.time);
            }
        }

        void SendSkeleton2(int sensorId, int user, Skeleton s, long time)
        {
            if (!capturing) { return; }
            
            if (!fullBody)
            {
                ProcessJointInformation(sensorId, user, 15, s.Joints[JointType.HandRight], s.BoneOrientations[JointType.HandRight], time);
            }
            else
            {
                ProcessJointInformation(sensorId, user, 1, s.Joints[JointType.Head], s.BoneOrientations[JointType.Head], time);
                ProcessJointInformation(sensorId, user, 2, s.Joints[JointType.ShoulderCenter], s.BoneOrientations[JointType.ShoulderCenter], time);
                ProcessJointInformation(sensorId, user, 3, s.Joints[JointType.Spine], s.BoneOrientations[JointType.Spine], time);
                ProcessJointInformation(sensorId, user, 4, s.Joints[JointType.HipCenter], s.BoneOrientations[JointType.HipCenter], time);
                // ProcessJointInformation(sensorId, user, 5, s.Joints[JointType.], s.BoneOrientations[JointType.], time);
                ProcessJointInformation(sensorId, user, 6, s.Joints[JointType.ShoulderLeft], s.BoneOrientations[JointType.ShoulderLeft], time);
                ProcessJointInformation(sensorId, user, 7, s.Joints[JointType.ElbowLeft], s.BoneOrientations[JointType.ElbowLeft], time);
                ProcessJointInformation(sensorId, user, 8, s.Joints[JointType.WristLeft], s.BoneOrientations[JointType.WristLeft], time);
                ProcessJointInformation(sensorId, user, 9, s.Joints[JointType.HandLeft], s.BoneOrientations[JointType.HandLeft], time);
                // ProcessJointInformation(sensorId, user, 10, s.Joints[JointType.], s.BoneOrientations[JointType.], time);
                // ProcessJointInformation(sensorId, user, 11, s.Joints[JointType.], s.BoneOrientations[JointType.], time);
                ProcessJointInformation(sensorId, user, 12, s.Joints[JointType.ShoulderRight], s.BoneOrientations[JointType.ShoulderRight], time);
                ProcessJointInformation(sensorId, user, 13, s.Joints[JointType.ElbowRight], s.BoneOrientations[JointType.ElbowRight], time);
                ProcessJointInformation(sensorId, user, 14, s.Joints[JointType.WristRight], s.BoneOrientations[JointType.WristRight], time);
                ProcessJointInformation(sensorId, user, 15, s.Joints[JointType.HandRight], s.BoneOrientations[JointType.HandRight], time);
                // ProcessJointInformation(sensorId, user, 16, s.Joints[JointType.], s.BoneOrientations[JointType.], time);
                ProcessJointInformation(sensorId, user, 17, s.Joints[JointType.HipLeft], s.BoneOrientations[JointType.HipLeft], time);
                ProcessJointInformation(sensorId, user, 18, s.Joints[JointType.KneeLeft], s.BoneOrientations[JointType.KneeLeft], time);
                ProcessJointInformation(sensorId, user, 19, s.Joints[JointType.AnkleLeft], s.BoneOrientations[JointType.AnkleLeft], time);
                ProcessJointInformation(sensorId, user, 20, s.Joints[JointType.FootLeft], s.BoneOrientations[JointType.FootLeft], time);
                ProcessJointInformation(sensorId, user, 21, s.Joints[JointType.HipRight], s.BoneOrientations[JointType.HipRight], time);
                ProcessJointInformation(sensorId, user, 22, s.Joints[JointType.KneeRight], s.BoneOrientations[JointType.KneeRight], time);
                ProcessJointInformation(sensorId, user, 23, s.Joints[JointType.AnkleRight], s.BoneOrientations[JointType.AnkleRight], time);
                ProcessJointInformation(sensorId, user, 24, s.Joints[JointType.FootRight], s.BoneOrientations[JointType.FootRight], stopwatch.ElapsedMilliseconds);
            }
        }

        double JointToConfidenceValue(Joint j)
        {
            if (j.TrackingState == JointTrackingState.Tracked) return 1;
            if (j.TrackingState == JointTrackingState.Inferred) return 0.5;
            if (j.TrackingState == JointTrackingState.NotTracked) return 0.1;
            return 0.5;
        }

        void ProcessJointInformation(int sensorId, int user, int joint, Joint j, BoneOrientation bo, long time)
        {
            SendJointMessage(sensorId, user, joint,
                j.Position.X, j.Position.Y, j.Position.Z,
                JointToConfidenceValue(j),
                time);
        }

        void SendJointMessage(int sensorId, int user, int joint, double x, double y, double z, double confidence, double time)
        {
            if (osc != null)
            {
                osc.Send(new OscElement("/joint", oscMapping[joint], sensorId, user, (float)(x * pointScale), (float)(-y * pointScale), (float)(z * pointScale), (float)confidence, time));
            }
            if (fileWriter != null)
            {
                // Joint, user, joint, x, y, z, on
                fileWriter.WriteLine("Joint," + sensorId + "," + user + "," + joint + "," +
                    (x * pointScale).ToString().Replace(",", ".") + "," +
                    (-y * pointScale).ToString().Replace(",", ".") + "," +
                    (z * pointScale).ToString().Replace(",", ".") + "," +
                    confidence.ToString().Replace(",", ".") + "," +
                    time.ToString().Replace(",", "."));
            }
        }

        void SendFacePoseMessage(int sensorId, int user, float x, float y, float z, float rotationX, float rotationY, float rotationZ, double time)
        {
            if (osc != null)
            {
                osc.Send(new OscElement(
                    "/face",
                    sensorId, user,
                    (float)(x * pointScale), (float)(-y * pointScale), (float)(z * pointScale),
                    rotationX, rotationY, rotationZ,
                    time));
            }
            if (fileWriter != null)
            {
                fileWriter.WriteLine("Face," +
                    sensorId + "," + user + "," +
                    (x * pointScale).ToString().Replace(",", ".") + "," +
                    (-y * pointScale).ToString().Replace(",", ".") + "," +
                    (z * pointScale).ToString().Replace(",", ".") + "," +
                    rotationX.ToString().Replace(",", ".") + "," +
                    rotationY.ToString().Replace(",", ".") + "," +
                    rotationZ.ToString().Replace(",", ".") + "," +
                    time.ToString().Replace(",", "."));
            }
        }

        void SendFaceAnimationMessage(int sensorId, int user, EnumIndexableCollection<AnimationUnit, float> c, double time)
        {
            if (osc != null)
            {
                osc.Send(new OscElement(
                    "/face_animation",
                    sensorId, user,
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
                    sensorId + "," + user + "," +
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
            List<KinectSensor> kinects = new List<KinectSensor>(sensors);
            foreach (KinectSensor kinect in kinects)
            {
                if (kinect == null) return;
                if (this.checkBoxSpeechCommands.IsChecked.Value)
                {
                    CreateSpeechRecognizer(kinect);
                }
                else
                {
                    StopSpeechRecognizer(kinect);
                }
            }
        }

        private void ImgClickSwitchKinect(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sensors.Count < 2) return;
            KinectSensor sensor = this.sensors.FirstOrDefault();
            if (sensor != null)
            {
                this.sensors.Remove(sensor);
                this.sensors.Add(sensor);
                SetFrames(sensor);
                SetFrames(sensors.FirstOrDefault());
            }
        }

        private void cbxDisplay_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            this.showRGB = this.cbxDisplay.SelectedIndex == 1;
            this.showDepth = this.cbxDisplay.SelectedIndex == 2;
            SetFrames(sensors.FirstOrDefault());
        }

    }
}