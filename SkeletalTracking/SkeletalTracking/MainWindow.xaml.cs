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
using Microsoft.Kinect.Toolkit.FaceTracking;
using Coding4Fun.Kinect.Wpf;
using System.IO;
using System.Diagnostics;
// Speech
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;
// OSC
using Ventuz.OSC;

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

        Byte[] colorPixelData;
        short[] depthPixelData;
        Skeleton[] allSkeletons;
        private SpeechRecognitionEngine speechRecognizer;
        private KinectSensor sensor;
        private Dictionary<int, FaceTracker> faceTrackers;
        private DispatcherTimer readyTimer;
        private UdpWriter osc;
        private StreamWriter fileWriter;
        private Stopwatch stopwatch;

        private const int skeletonCount = 6;
        private double pointScale = 1000;
        private bool shuttingDown = false;

        private bool allUsers = true;
        private bool fullBody = true;
        private bool faceTracking = true;
        private bool faceTracking2DMesh = false;
        private bool faceTrackingHeadPose = true;
        private bool faceTrackingAnimationUnits = true;
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

            CheckForShortcut();
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
                fileWriter.WriteLine("Joint, user, joint, x, y, z, confidence, time");
                fileWriter.WriteLine("Face, user, x, y, z, pitch, yaw, roll, time");
                fileWriter.WriteLine("FaceAnimation, face, lip_raise, lip_stretcher, lip_corner_depressor, jaw_lower, brow_lower, brow_raise, time");

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
                Smoothing = 0.1f,
                Correction = 0.1f,
                Prediction = 0.1f,
                JitterRadius = 0.01f,
                MaxDeviationRadius = 0.04f
            };

            if (faceTracking)
            {
                sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
                sensor.DepthStream.Range = DepthRange.Default;
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                sensor.SkeletonStream.Enable(parameters);
                colorPixelData = new byte[sensor.ColorStream.FramePixelDataLength];
                depthPixelData = new short[sensor.DepthStream.FramePixelDataLength];
                faceTrackers = new Dictionary<int, FaceTracker>();
                allSkeletons = new Skeleton[skeletonCount];
                //FaceTracker f = new FaceTracker(sensor);
                sensor.AllFramesReady += new EventHandler<AllFramesReadyEventArgs>(sensor_AllFramesReady);
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
                allSkeletons = new Skeleton[skeletonCount];
                sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);
            }

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

        void sensor_AllFramesReady(object sender, AllFramesReadyEventArgs e)
        {

            if (shuttingDown)
            {
                return;
            }

            ColorImageFrame colorFrame = e.OpenColorImageFrame();
            if (colorFrame == null) return;
            colorFrame.CopyPixelDataTo(colorPixelData);
            DepthImageFrame depthFrame = e.OpenDepthImageFrame();
            if (depthFrame == null) return;
            depthFrame.CopyPixelDataTo(depthPixelData);
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
                    SendFaceTracking(0, first);
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
                        SendFaceTracking(s.TrackingId, s);
                    }
                }
            }

            ScalePosition2(headImage, first.Joints[JointType.Head]);
            ScalePosition2(leftEllipse, first.Joints[JointType.HandLeft]);
            ScalePosition2(rightEllipse, first.Joints[JointType.HandRight]);
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

        void SendFaceTracking(int user, Skeleton s)
        {
            if (!faceTrackers.ContainsKey(user))
            {
                try
                {
                    faceTrackers.Add(user, new FaceTracker(sensor));
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

            FaceTrackFrame faceFrame = faceTrackers[user].Track(
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
                        faceFrame.Rotation.X, faceFrame.Rotation.Y, faceFrame.Rotation.Z, stopwatch.ElapsedMilliseconds);
                }

                if (faceTrackingAnimationUnits)
                {
                    SendFaceAnimationMessage(user, faceFrame.GetAnimationUnitCoefficients(), stopwatch.ElapsedMilliseconds);
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
                Status.Content = -y * pointScale;
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

        private void StopKinect(KinectSensor sensor)
        {
            if (sensor != null)
            {
                if (sensor.IsRunning)
                {

                    sensor.ColorStream.Disable();
                    sensor.DepthStream.Disable();
                    sensor.SkeletonStream.Disable();

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
                        MessageBox.Show("Missing company or description: " + company + " - " + description);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetErrorText(ex));
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
