using Microsoft.Kinect;
using Microsoft.Kinect.Toolkit.FaceTracking;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Ventuz.OSC;

namespace Microsoft.Samples.Kinect.SkeletonBasics
{
    abstract class TrackingInformation
    {
        public int sensorId;
        public abstract void Send(UdpWriter osc, StreamWriter fileWriter, int pointScale);
    }

    class SkeletonTrackingInformation : TrackingInformation {

        public int user;
        public Skeleton skeleton;
        public bool fullBody;
        public double time;

        public static List<String> oscMapping = new List<String> { "",
            "head", "neck", "torso", "waist",
            "l_collar", "l_shoulder", "l_elbow", "l_wrist", "l_hand", "l_fingertip",
            "r_collar", "r_shoulder", "r_elbow", "r_wrist", "r_hand", "r_fingertip",
            "l_hip", "l_knee", "l_ankle", "l_foot",
            "r_hip", "r_knee", "r_ankle", "r_foot" };

        public SkeletonTrackingInformation(int sensorId, int user, Skeleton skeleton, bool fullBody, double time)
        {
            this.sensorId = sensorId;
            this.user = user;
            this.skeleton = skeleton;
            this.fullBody = fullBody;
            this.time = time;
        }

        public override void Send(UdpWriter osc, StreamWriter fileWriter, int pointScale)
        {
            if (!fullBody)
            {
                ProcessJointInformation(15, skeleton.Joints[JointType.HandRight], skeleton.BoneOrientations[JointType.HandRight], time, osc, fileWriter, pointScale);
            }
            else
            {
                ProcessJointInformation(1, skeleton.Joints[JointType.Head], skeleton.BoneOrientations[JointType.Head], time, osc, fileWriter, pointScale);
                ProcessJointInformation(2, skeleton.Joints[JointType.ShoulderCenter], skeleton.BoneOrientations[JointType.ShoulderCenter], time, osc, fileWriter, pointScale);
                ProcessJointInformation(3, skeleton.Joints[JointType.Spine], skeleton.BoneOrientations[JointType.Spine], time, osc, fileWriter, pointScale);
                ProcessJointInformation(4, skeleton.Joints[JointType.HipCenter], skeleton.BoneOrientations[JointType.HipCenter], time, osc, fileWriter, pointScale);
                // ProcessJointInformation(5, skeleton.Joints[JointType.], skeleton.BoneOrientations[JointType.], time, osc, fileWriter, pointScale);
                ProcessJointInformation(6, skeleton.Joints[JointType.ShoulderLeft], skeleton.BoneOrientations[JointType.ShoulderLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(7, skeleton.Joints[JointType.ElbowLeft], skeleton.BoneOrientations[JointType.ElbowLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(8, skeleton.Joints[JointType.WristLeft], skeleton.BoneOrientations[JointType.WristLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(9, skeleton.Joints[JointType.HandLeft], skeleton.BoneOrientations[JointType.HandLeft], time, osc, fileWriter, pointScale);
                // ProcessJointInformation(10, skeleton.Joints[JointType.], skeleton.BoneOrientations[JointType.], time, osc, fileWriter, pointScale);
                // ProcessJointInformation(11, skeleton.Joints[JointType.], skeleton.BoneOrientations[JointType.], time, osc, fileWriter, pointScale);
                ProcessJointInformation(12, skeleton.Joints[JointType.ShoulderRight], skeleton.BoneOrientations[JointType.ShoulderRight], time, osc, fileWriter, pointScale);
                ProcessJointInformation(13, skeleton.Joints[JointType.ElbowRight], skeleton.BoneOrientations[JointType.ElbowRight], time, osc, fileWriter, pointScale);
                ProcessJointInformation(14, skeleton.Joints[JointType.WristRight], skeleton.BoneOrientations[JointType.WristRight], time, osc, fileWriter, pointScale);
                ProcessJointInformation(15, skeleton.Joints[JointType.HandRight], skeleton.BoneOrientations[JointType.HandRight], time, osc, fileWriter, pointScale);
                // ProcessJointInformation(16, skeleton.Joints[JointType.], skeleton.BoneOrientations[JointType.], time, osc, fileWriter, pointScale);
                ProcessJointInformation(17, skeleton.Joints[JointType.HipLeft], skeleton.BoneOrientations[JointType.HipLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(18, skeleton.Joints[JointType.KneeLeft], skeleton.BoneOrientations[JointType.KneeLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(19, skeleton.Joints[JointType.AnkleLeft], skeleton.BoneOrientations[JointType.AnkleLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(20, skeleton.Joints[JointType.FootLeft], skeleton.BoneOrientations[JointType.FootLeft], time, osc, fileWriter, pointScale);
                ProcessJointInformation(21, skeleton.Joints[JointType.HipRight], skeleton.BoneOrientations[JointType.HipRight], time, osc, fileWriter, pointScale);
                ProcessJointInformation(22, skeleton.Joints[JointType.KneeRight], skeleton.BoneOrientations[JointType.KneeRight], time, osc, fileWriter, pointScale);
                ProcessJointInformation(23, skeleton.Joints[JointType.AnkleRight], skeleton.BoneOrientations[JointType.AnkleRight], time, osc, fileWriter, pointScale);
                ProcessJointInformation(24, skeleton.Joints[JointType.FootRight], skeleton.BoneOrientations[JointType.FootRight], time, osc, fileWriter, pointScale);
            }
        }


        double JointToConfidenceValue(Joint j)
        {
            if (j.TrackingState == JointTrackingState.Tracked) return 1;
            if (j.TrackingState == JointTrackingState.Inferred) return 0.5;
            if (j.TrackingState == JointTrackingState.NotTracked) return 0.1;
            return 0.5;
        }

        void ProcessJointInformation(int joint, Joint j, BoneOrientation bo, double time, UdpWriter osc, StreamWriter fileWriter, int pointScale)
        {
            SendJointMessage(joint,
                j.Position.X, j.Position.Y, j.Position.Z,
                JointToConfidenceValue(j), time,
                osc, fileWriter, pointScale);
        }

        void SendJointMessage(int joint, double x, double y, double z, double confidence, double time, UdpWriter osc, StreamWriter fileWriter, int pointScale)
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
    }

    class FacePoseTrackingInformation : TrackingInformation {
        public int user;
        public float x, y, z;
        public float rotationX, rotationY, rotationZ;
        public double time;

        public FacePoseTrackingInformation(int sensorId, int user, float x, float y, float z, float rotationX, float rotationY, float rotationZ, double time)
        {
            this.sensorId = sensorId;
            this.user = user;
            this.x = x;
            this.y = y;
            this.z = z;
            this.rotationX = rotationX;
            this.rotationY = rotationY;
            this.rotationZ = rotationZ;
            this.time = time;
        }

        public override void Send(UdpWriter osc, StreamWriter fileWriter, int pointScale)
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
    }
    
    class FaceAnimationTrackingInformation : TrackingInformation {
            int user;
            double time;
            EnumIndexableCollection<AnimationUnit, float> c;

            public FaceAnimationTrackingInformation(int sensorId, int user, EnumIndexableCollection<AnimationUnit, float> c, double time)
            {
                this.sensorId = sensorId;
                this.user = user;
                this.c = c;
                this.time = time;
            }

            public override void Send(UdpWriter osc, StreamWriter fileWriter, int pointScale)
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
    }
}
