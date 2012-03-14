// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Public License (Ms-PL).
// Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Coding4Fun.Kinect.Wpf;
using System.IO; 
//speech
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
//OSC
using Ventuz.OSC;
using System.Diagnostics;


namespace SkeletalTracking
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        bool closing = false;
        const int skeletonCount = 6; 
        Skeleton[] allSkeletons = new Skeleton[skeletonCount];
        private SpeechRecognitionEngine speechRecognizer;
        private bool running = true;
        private KinectSensor sensor;
        private DispatcherTimer readyTimer;
        private UdpWriter osc;
        private StreamWriter sw;
        private Stopwatch stopwatch;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.osc = new UdpWriter("127.0.0.1", 7110);
            //test: send 4 points
            osc.Send(new OscElement("/joint", "r_hand", 0, 0, 0, 1000));
            osc.Send(new OscElement("/joint", "r_hand", 0, -350, -400, 1500));
            osc.Send(new OscElement("/joint", "r_hand", 0, -350, 0, 2000));
            osc.Send(new OscElement("/joint", "r_hand", 0, 0, -400, 2500));
            //create csv file
            stopwatch = new Stopwatch();
            stopwatch.Start();
            sw = new StreamWriter("points.csv", true);
            sw.WriteLine("Joint, type, id, x, y, on");
            kinectSensorChooser1.KinectSensorChanged += new DependencyPropertyChangedEventHandler(kinectSensorChooser1_KinectSensorChanged);
            //this.kinectColorViewer1.Visibility = System.Windows.Visibility.Hidden;

        }

        void kinectSensorChooser1_KinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            KinectSensor old = (KinectSensor)e.OldValue;

            StopKinect(old);

            //KinectSensor sensor = (KinectSensor)e.NewValue;
            sensor = (KinectSensor)e.NewValue;

            if (sensor == null)
            {
                return;
            }

            this.osc = new UdpWriter("127.0.0.1", 7110);
            //this.osc = new UdpWriter("192.168.6.144", 7110);

            var parameters = new TransformSmoothParameters
            {
                Smoothing = 0.3f,
                Correction = 0.2f,
                Prediction = 0.0f,
                JitterRadius = 1.0f,
                MaxDeviationRadius = 0.5f
            };
            //sensor.SkeletonStream.Enable(parameters);

            sensor.SkeletonStream.Enable();

            sensor.AllFramesReady += new EventHandler<AllFramesReadyEventArgs>(sensor_AllFramesReady);
            sensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30); 
            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);

            this.speechRecognizer = this.CreateSpeechRecognizer();

            try
            {
                sensor.Start();
            }
            catch (System.IO.IOException)
            {
                kinectSensorChooser1.AppConflictOccurred();
            }

            if (this.speechRecognizer != null && sensor != null)
            {
                // NOTE: Need to wait 4 seconds for device to be ready to stream audio right after initialization
                this.readyTimer = new DispatcherTimer();
                this.readyTimer.Tick += this.ReadyTimerTick;
                this.readyTimer.Interval = new TimeSpan(0, 0, 4);
                this.readyTimer.Start();

                Status.Content = "Init";
                //this.ReportSpeechStatus("Initializing audio stream...");
                //this.UpdateInstructionsText(string.Empty);

                //this.Closing += this.MainWindowClosing;
            }
            else
            {
                Status.Content = "FAILED";
            }
        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.Start();
            //this.ReportSpeechStatus("Ready to recognize speech!");
            //this.UpdateInstructionsText("Say: 'red', 'green' or 'blue'");
            Status.Content = "STOP";
            this.readyTimer.Stop();
            this.readyTimer = null;
        }

        void sensor_AllFramesReady(object sender, AllFramesReadyEventArgs e)
        {
            if (closing)
            {
                return;
            }

            //Get a skeleton
            Skeleton first =  GetFirstSkeleton(e);

            if (first == null)
            {
                return; 
            }



            //set scaled position
            //ScalePosition(headImage, first.Joints[JointType.Head]);
            ScalePosition(leftEllipse, first.Joints[JointType.HandLeft]);
            ScalePosition(rightEllipse, first.Joints[JointType.HandRight]);

            GetCameraPoint(first, e); 

        }

        void GetCameraPoint(Skeleton first, AllFramesReadyEventArgs e)
        {

            using (DepthImageFrame depth = e.OpenDepthImageFrame())
            {
                if (depth == null ||
                    kinectSensorChooser1.Kinect == null)
                {
                    return;
                }
                

                //Map a joint location to a point on the depth map
                //head
                DepthImagePoint headDepthPoint =
                    depth.MapFromSkeletonPoint(first.Joints[JointType.Head].Position);
                //left hand
                DepthImagePoint leftDepthPoint =
                    depth.MapFromSkeletonPoint(first.Joints[JointType.HandLeft].Position);
                //right hand
                DepthImagePoint rightDepthPoint =
                    depth.MapFromSkeletonPoint(first.Joints[JointType.HandRight].Position);


                //Map a depth point to a point on the color image
                //head
                ColorImagePoint headColorPoint =
                    depth.MapToColorImagePoint(headDepthPoint.X, headDepthPoint.Y,
                    ColorImageFormat.RgbResolution640x480Fps30);
                //left hand
                ColorImagePoint leftColorPoint =
                    depth.MapToColorImagePoint(leftDepthPoint.X, leftDepthPoint.Y,
                    ColorImageFormat.RgbResolution640x480Fps30);
                //right hand
                ColorImagePoint rightColorPoint =
                    depth.MapToColorImagePoint(rightDepthPoint.X, rightDepthPoint.Y,
                    ColorImageFormat.RgbResolution640x480Fps30);

                //zet hier osc
                if (running == true)
                {
                    SkeletonPoint ShoulderLeft = first.Joints[JointType.ShoulderLeft].Position;
                    SkeletonPoint ShoulderRight = first.Joints[JointType.ShoulderRight].Position;
                    SkeletonPoint HandLeft = first.Joints[JointType.HandLeft].Position;
                    SkeletonPoint HandRight = first.Joints[JointType.HandRight].Position;
                    //osc.Send(new OscElement("/joint", "l_shoulder", 0, 1000 * ShoulderLeft.X, -1000 * ShoulderLeft.Y, 1000 * ShoulderLeft.Z));
                    //osc.Send(new OscElement("/joint", "r_shoulder", 0, 1000 * ShoulderRight.X, -1000 * ShoulderRight.Y, 1000 * ShoulderRight.Z));
                    osc.Send(new OscElement("/joint", "r_hand", 0, 1000 * HandRight.X, 1000 * HandRight.Y, 1000 * HandRight.Z));

                    sw.WriteLine("r_hand, 0, " + 1000 * HandRight.X + ", " + 1000 * HandRight.Y + ", " + stopwatch.ElapsedMilliseconds);
                    //osc.Send(new OscElement("/joint", "l_hand", 0, 1000 * HandLeft.X, -1000 * HandLeft.Y, 1000 * HandLeft.Z));
                    Status.Foreground = Brushes.Black;
                    /*osc.Send(new OscElement("/joint", "r_hand", 0, 350, 200, 300));
                    osc.Send(new OscElement("/joint", "r_hand", 0, 0, -200, 300));
                    osc.Send(new OscElement("/joint", "r_hand", 0, 0, 200, 300));
                    osc.Send(new OscElement("/joint", "r_hand", 0, 350, -200, 300));*/
                    if (Distance2D(1000 * HandRight.X, 1000 * HandRight.Y, 0, 0) < 150)
                    {
                        Status.Content = 1;
                    }
                    else if (Distance2D(1000 * HandRight.X, 1000 * HandRight.Y, -350, -400) < 150)
                    {
                        Status.Content = 2;
                    }
                    else if (Distance2D(1000 * HandRight.X, 1000 * HandRight.Y, -350, 0) < 150)
                    {
                        Status.Content = 3;
                    }
                    else if (Distance2D(1000 * HandRight.X, 1000 * HandRight.Y, 0, -400) < 150)
                    {
                        Status.Content = 4;
                    }
                    else { Status.Content = 0; }
                    //Status.Content = (int)(1000 * HandRight.X) + "/" + (int)(-1000 * HandRight.Y);
                }

                //Set location
                CameraPosition(headImage, headColorPoint);
                CameraPosition(leftEllipse, leftColorPoint);
                CameraPosition(rightEllipse, rightColorPoint);
            }        
        }

        private int Distance2D(double x1, double y1, int x2, int y2)
        {
            //     ______________________
            //d = &#8730; (x2-x1)^2 + (y2-y1)^2
            //

            //Our end result
            int result = 0;
            //Take x2-x1, then square it
            double part1 = Math.Pow((x2 - x1), 2);
            //Take y2-y1, then sqaure it
            double part2 = Math.Pow((y2 - y1), 2);
            //Add both of the parts together
            double underRadical = part1 + part2;
            //Get the square root of the parts
            result = (int)Math.Sqrt(underRadical);
            //Return our result
            return result;
        }


        Skeleton GetFirstSkeleton(AllFramesReadyEventArgs e)
        {
            using (SkeletonFrame skeletonFrameData = e.OpenSkeletonFrame())
            {
                if (skeletonFrameData == null)
                {
                    return null; 
                }

                
                skeletonFrameData.CopySkeletonDataTo(allSkeletons);

                //get the first tracked skeleton
                Skeleton first = (from s in allSkeletons
                                         where s.TrackingState == SkeletonTrackingState.Tracked
                                         select s).FirstOrDefault();

                return first;

            }
        }

        private void StopKinect(KinectSensor sensor)
        {
            if (sensor != null)
            {
                if (sensor.IsRunning)
                {
                    //stop sensor 
                    sensor.Stop();

                    //stop audio if not null
                    if (sensor.AudioSource != null)
                    {
                        sensor.AudioSource.Stop();
                    }


                }
            }
        }

        private void CameraPosition(FrameworkElement element, ColorImagePoint point)
        {
            //Divide by 2 for width and height so point is right in the middle 
            // instead of in top/left corner
            Canvas.SetLeft(element, point.X - element.Width / 2);
            Canvas.SetTop(element, point.Y - element.Height / 2);

        }

        private void ScalePosition(FrameworkElement element, Joint joint)
        {
            //convert the value to X/Y
            //Joint scaledJoint = joint.ScaleTo(1280, 720); 
            
            //convert & scale (.3 = means 1/3 of joint distance)
            Joint scaledJoint = joint.ScaleTo(1280, 720, .3f, .3f);

            Canvas.SetLeft(element, scaledJoint.Position.X);
            Canvas.SetTop(element, scaledJoint.Position.Y); 
            
        }

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        private SpeechRecognitionEngine CreateSpeechRecognizer()
        {
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Close();
                return null;
            }

            //SpeechRecognitionEngine sre;
            try
            {
                speechRecognizer = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                this.Close();
                return null;
            }

            var grammar = new Choices();
            grammar.Add("start");
            grammar.Add("stop");

            var gb = new GrammarBuilder { Culture = ri.Culture };
            gb.Append(grammar);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            speechRecognizer.LoadGrammar(g);
            speechRecognizer.SpeechRecognized += this.SreSpeechRecognized;
            speechRecognizer.SpeechRecognitionRejected += this.SreSpeechRecognitionRejected;

            return speechRecognizer;
        }

        private void Start()
        {
            var audioSource = this.sensor.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();
            //this.stream = new EnergyCalculatingPassThroughStream(kinectStream);
            this.speechRecognizer.SetInputToAudioStream(
                kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            this.speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
            //var t = new Thread(this.PollSoundSourceLocalization);
            //t.Start();
        }

        private void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            //this.RejectSpeech(e.Result);
            Status.Foreground = Brushes.Black;
            Status.Content = "R: " + e.Result.Text;
        }

        private void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Confidence < 0.05)
            {
                Status.Foreground = Brushes.Black;
                Status.Content = "L: " + e.Result.Text;
                return;
            }

            switch (e.Result.Text)
            {
                case "start":
                    //brush = this.redBrush;
                    Status.Foreground = Brushes.Green;
                    Status.Content = "START";
                    running = true;
                    break;
                case "stop":
                    Status.Foreground = Brushes.Red;
                    Status.Content = "STOP";
                    running = false;
                    //brush = this.greenBrush;
                    break;
                default:
                    //brush = this.blackBrush;
                    Status.Content = e.Result.Text;
                    break;
            }
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            closing = true; 
            StopKinect(kinectSensorChooser1.Kinect);
            sw.Close();
            stopwatch.Stop();
        }
    }
}
