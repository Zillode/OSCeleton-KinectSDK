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

        Skeleton[] allSkeletons = new Skeleton[skeletonCount];
        private SpeechRecognitionEngine speechRecognizer;
        private KinectSensor sensor;
        private DispatcherTimer readyTimer;
        private UdpWriter osc;
        private StreamWriter fileWriter;
        private Stopwatch stopwatch;

        private const int skeletonCount = 6;
        private double pointScale = 1000;
        private bool shuttingDown = false;

        private bool allUsers = true;
        private bool fullBody = true;
        private bool capturing = true;
        private bool writeOSC = true;
        private bool writeFile = true;
        private bool voiceRecognition = false;

        private static List<String> oscMapping = new List<String> { "",
            "head", "neck", "torso", "waist",
            "l_collar", "l_shoulder", "l_elbow", "l_wrist", "l_hand", "l_fingertip",
            "r_collar", "r_shoulder", "r_elbow", "r_wrist", "r_hand", "r_fingertip",
            "l_hip", "l_knee", "l_ankle", "l_foot",
            "r_hip", "r_knee", "r_ankle", "r_foot" };


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            stopwatch = new Stopwatch();
            stopwatch.Reset();
            stopwatch.Start();
            if (writeOSC)
            {
                osc = new UdpWriter("127.0.0.1", 7110);
            }
            if (writeFile)
            {
                fileWriter = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "/" + string.Format("points-{0:yyyy-MM-dd_hh-mm-ss-tt}.csv", DateTime.Now), false);
                fileWriter.WriteLine("Joint, user, joint, x, y, z, on");
            }
            kinectSensorChooser1.KinectSensorChanged += new DependencyPropertyChangedEventHandler(kinectSensorChooser1_KinectSensorChanged);
        }

        void kinectSensorChooser1_KinectSensorChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            KinectSensor old = (KinectSensor)e.OldValue;
            StopKinect(old);
            sensor = (KinectSensor)e.NewValue;
            if (sensor == null)
            {
                return;
            }

            var parameters = new TransformSmoothParameters
            {
                Smoothing = 0.02f,
                Correction = 0.1f,
                Prediction = 0.1f,
                JitterRadius = 0.05f,
                MaxDeviationRadius = 0.04f
            };

            sensor.SkeletonStream.Enable(parameters);
            sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);

            if (voiceRecognition)
            {
                this.speechRecognizer = this.CreateSpeechRecognizer();
            }

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
                // Wait 4 seconds for device to be ready to stream audio
                this.readyTimer = new DispatcherTimer();
                this.readyTimer.Tick += this.ReadyTimerTick;
                this.readyTimer.Interval = new TimeSpan(0, 0, 4);
                this.readyTimer.Start();

                Status.Content = "Initialising";
            }
        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.Start();
            Status.Content = "Listening...";
            this.readyTimer.Stop();
            this.readyTimer = null;
        }

        void sensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {

            if (shuttingDown)
            {
                return;
            }

            SkeletonFrame skeletonFrame = e.OpenSkeletonFrame();
            if (skeletonFrame == null) return;
            skeletonFrame.CopySkeletonDataTo(allSkeletons);

            var first = (from s in allSkeletons
                         where s.TrackingState == SkeletonTrackingState.Tracked
                         select s).FirstOrDefault();
            if (first == null) return;

            if (capturing)
            {
                if (!allUsers)
                {
                        SendSkeleton(0, first);
                }
                else
                {
                    IEnumerable<Skeleton> trackedSkeletons = (
                            from s in allSkeletons
                            where s.TrackingState == SkeletonTrackingState.Tracked
                            select s);
                    foreach (Skeleton s in trackedSkeletons)
                    {
                        SendSkeleton(s.TrackingId, s);
                    }
                }
            }

            ScalePosition2(headImage, first.Joints[JointType.Head]);
            ScalePosition2(leftEllipse, first.Joints[JointType.HandLeft]);
            ScalePosition2(rightEllipse, first.Joints[JointType.HandRight]);
        }

        void SendSkeleton(int user, Skeleton s)
        {
            if (!capturing) { return; }

            if (!fullBody)
            {
                SendMessage(user, 15, s.Joints[JointType.HandRight].Position.X, s.Joints[JointType.HandRight].Position.Y, s.Joints[JointType.HandRight].Position.Z);
            }
            else
            {
                SendMessage(user, 1, s.Joints[JointType.Head].Position.X, s.Joints[JointType.Head].Position.Y, s.Joints[JointType.Head].Position.Z);
                SendMessage(user, 2, s.Joints[JointType.ShoulderCenter].Position.X, s.Joints[JointType.ShoulderCenter].Position.Y, s.Joints[JointType.ShoulderCenter].Position.Z);
                SendMessage(user, 3, s.Joints[JointType.Spine].Position.X, s.Joints[JointType.Spine].Position.Y, s.Joints[JointType.Spine].Position.Z);
                SendMessage(user, 4, s.Joints[JointType.HipCenter].Position.X, s.Joints[JointType.HipCenter].Position.Y, s.Joints[JointType.HipCenter].Position.Z);
                // SendMessage(user, 5, s.Joints[JointType.].Position.X, s.Joints[JointType.].Position.Y, s.Joints[JointType.].Position.Z);
                SendMessage(user, 6, s.Joints[JointType.ShoulderLeft].Position.X, s.Joints[JointType.ShoulderLeft].Position.Y, s.Joints[JointType.ShoulderLeft].Position.Z);
                SendMessage(user, 7, s.Joints[JointType.ElbowLeft].Position.X, s.Joints[JointType.ElbowLeft].Position.Y, s.Joints[JointType.ElbowLeft].Position.Z);
                SendMessage(user, 8, s.Joints[JointType.WristLeft].Position.X, s.Joints[JointType.WristLeft].Position.Y, s.Joints[JointType.WristLeft].Position.Z);
                SendMessage(user, 9, s.Joints[JointType.HandLeft].Position.X, s.Joints[JointType.HandLeft].Position.Y, s.Joints[JointType.HandLeft].Position.Z);
                // SendMessage(user, 10, s.Joints[JointType.].Position.X, s.Joints[JointType.].Position.Y, s.Joints[JointType.].Position.Z);
                // SendMessage(user, 11, s.Joints[JointType.].Position.X, s.Joints[JointType.].Position.Y, s.Joints[JointType.].Position.Z);
                SendMessage(user, 12, s.Joints[JointType.ShoulderRight].Position.X, s.Joints[JointType.ShoulderRight].Position.Y, s.Joints[JointType.ShoulderRight].Position.Z);
                SendMessage(user, 13, s.Joints[JointType.ElbowRight].Position.X, s.Joints[JointType.ElbowRight].Position.Y, s.Joints[JointType.ElbowRight].Position.Z);
                SendMessage(user, 14, s.Joints[JointType.WristRight].Position.X, s.Joints[JointType.WristRight].Position.Y, s.Joints[JointType.WristRight].Position.Z);
                SendMessage(user, 15, s.Joints[JointType.HandRight].Position.X, s.Joints[JointType.HandRight].Position.Y, s.Joints[JointType.HandRight].Position.Z);
                // SendMessage(user, 16, s.Joints[JointType.].Position.X, s.Joints[JointType.].Position.Y, s.Joints[JointType.].Position.Z);
                SendMessage(user, 17, s.Joints[JointType.HipLeft].Position.X, s.Joints[JointType.HipLeft].Position.Y, s.Joints[JointType.HipLeft].Position.Z);
                SendMessage(user, 18, s.Joints[JointType.KneeLeft].Position.X, s.Joints[JointType.KneeLeft].Position.Y, s.Joints[JointType.KneeLeft].Position.Z);
                SendMessage(user, 19, s.Joints[JointType.AnkleLeft].Position.X, s.Joints[JointType.AnkleLeft].Position.Y, s.Joints[JointType.AnkleLeft].Position.Z);
                SendMessage(user, 20, s.Joints[JointType.FootLeft].Position.X, s.Joints[JointType.FootLeft].Position.Y, s.Joints[JointType.FootLeft].Position.Z);
                SendMessage(user, 21, s.Joints[JointType.HipRight].Position.X, s.Joints[JointType.HipRight].Position.Y, s.Joints[JointType.HipRight].Position.Z);
                SendMessage(user, 22, s.Joints[JointType.KneeRight].Position.X, s.Joints[JointType.KneeRight].Position.Y, s.Joints[JointType.KneeRight].Position.Z);
                SendMessage(user, 23, s.Joints[JointType.AnkleRight].Position.X, s.Joints[JointType.AnkleRight].Position.Y, s.Joints[JointType.AnkleRight].Position.Z);
                SendMessage(user, 24, s.Joints[JointType.FootRight].Position.X, s.Joints[JointType.FootRight].Position.Y, s.Joints[JointType.FootRight].Position.Z);
            }
        }

        void SendMessage(int user, int joint, double x, double y, double z)
        {
            if (osc != null)
            {
                Status.Content = -y * pointScale;
                osc.Send(new OscElement("/joint", oscMapping[joint], user, (int)(x * pointScale), (int)(-y * pointScale), (int)(z * pointScale)));
            }
            if (fileWriter != null)
            {
                // Joint, user, joint, x, y, z, on
                fileWriter.WriteLine("Joint," + user + "," + joint + "," +
                    (x * pointScale).ToString().Replace(",", ".") + "," +
                    (-y * pointScale).ToString().Replace(",", ".") + "," +
                    (z * pointScale).ToString().Replace(",", ".") + "," +
                    stopwatch.ElapsedMilliseconds.ToString().Replace(",", "."));
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

        private void ScalePosition(FrameworkElement element, Joint joint)
        {
            //convert the value to X/Y
            //Joint scaledJoint = joint.ScaleTo(1280, 720);

            //convert & scale (.3 = means 1/3 of joint distance)
            Joint scaledJoint = joint.ScaleTo(1280, 720, .3f, .3f);

            Canvas.SetLeft(element, scaledJoint.Position.X);
            Canvas.SetTop(element, scaledJoint.Position.Y);

        }

        private void ScalePosition2(FrameworkElement element, Joint joint)
        {
            //convert the value to X/Y
            //Joint scaledJoint = joint.ScaleTo(1280, 720);

            //convert & scale (.3 = means 1/3 of joint distance)
            Joint scaledJoint = joint.ScaleTo(640, 480);//joint.ScaleTo(640, 480, .3f, .3f);

            Canvas.SetLeft(element, scaledJoint.Position.X - (element.Width / 2));
            Canvas.SetTop(element, scaledJoint.Position.Y - (element.Width / 2));

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
            grammar.Add("camera on");
            grammar.Add("camera off");

            var gb = new GrammarBuilder { Culture = ri.Culture };
            gb.Append(grammar);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            speechRecognizer.LoadGrammar(g);
            speechRecognizer.SpeechRecognized += this.SreSpeechRecognized;
            speechRecognizer.SpeechRecognitionRejected += this.SreSpeechRecognitionRejected;
            speechRecognizer.SpeechHypothesized += this.SreSpeechHypothesized;

            return speechRecognizer;
        }

        private void Start()
        {
            var audioSource = this.sensor.AudioSource;
            if (audioSource == null)
            {
                return;
            }
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();
            //this.stream = new EnergyCalculatingPassThroughStream(kinectStream);
            this.speechRecognizer.SetInputToAudioStream(
                kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            this.speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);
            //var t = new Thread(this.PollSoundSourceLocalization);
            //t.Start();
        }

        private void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            Status.Foreground = Brushes.Black;
            Status.Content = "H: " + e.Result.Text;
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
                Status.Content = "A: " + e.Result.Text;
                return;
            }

            switch (e.Result.Text.ToUpperInvariant())
            {
                case "START":
                    Status.Foreground = Brushes.Green;
                    Status.Content = "START";
                    capturing = true;
                    break;
                case "STOP":
                    Status.Foreground = Brushes.Red;
                    Status.Content = "STOP";
                    capturing = false;
                    break;
                case "CAMERA ON":
                    Status.Foreground = Brushes.Green;
                    Status.Content = "ON";
                    capturing = true;
                    break;
                case "CAMERA OFF":
                    Status.Foreground = Brushes.Red;
                    Status.Content = "OFF";
                    capturing = false;
                    break;
                default:
                    Status.Content = e.Result.Text;
                    break;
            }
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            shuttingDown = true;
            StopKinect(kinectSensorChooser1.Kinect);
            if (fileWriter != null)
            {
                fileWriter.Close();
                fileWriter = null;
            }
            stopwatch.Stop();
        }
    }
}
