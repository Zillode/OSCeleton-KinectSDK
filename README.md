OSCeleton-KinectSDK
===================

What is this?
-------------

As the title says, it's just a small program that takes kinect
skeleton data from the KinectSDK (v1.8) framework and spits out the coordinates
of the skeleton's joints via OSC messages. These can can then be used
on your language / framework of choice.

The OSC protocol of this application is compatible with the [OSCeleton-OpenNI](https://github.com/Zillode/OSCeleton-OpenNI).
A similar application for KinectSDK v2.0 is available [here](https://github.com/Zillode/OSCeleton-KinectSDK2).

Note: this protocol is partially incompatible with Sensebloom/OSCeleton!

How do I use it?
----------------

Download and install the [ClickOnce installer](http://osceleton.zillode.be/private/osceleton-kinectsdk/setup.exe)

How do I build it?
------------------

### Install Visual Studio 2012
### Install [Microsoft Kinect SDK (version 1.8)](https://www.microsoft.com/en-us/download/details.aspx?id=40278)
### Install [Microsoft Speech Platform SDK (version 11)](http://www.microsoft.com/en-us/download/details.aspx?id=27226)
### Compile and run the SkeletonBasics solution

If you run the executable, it will send the OSC
messages in the Midas format to localhost on port 7110.
To learn about the OSC message format, continue reading below.


OSC Message format
------------------

### Joint message - message with the coordinates of each skeleton joint:
The messages will have the following format:

    Address pattern: "/joint"
    Type tag: "siiffffd"
    s: Joint name, check out the full list of joints below
    i: The ID of the sensor
    i: The ID of the user
    f: X coordinate of joint in real world coordinates (centimers)
    f: Y coordinate of joint in real world coordinates (centimers)
    f: Z coordinate of joint in real world coordinates (centimers)
    f: confidence value in interval [0.0, 1.0]
	d: timestamp in milliseconds since launch

Note: the Y coordinate is inverted compared to the default KinectSDK to be compatible with OpenNI.

### Face message - message with the coordinates of a face event:
The messages will have the following format:

    Address pattern: "/face"
    Type tag: "iiffffd"
    i: The ID of the sensor
    i: The ID of the user
    f: X coordinate of joint in real world coordinates (centimers)
    f: Y coordinate of joint in real world coordinates (centimers)
    f: Z coordinate of joint in real world coordinates (centimers)
    f: pitch of the head [-90, 90]
    f: pitch of the yaw [-90, 90]
    f: pitch of the roll [-90, 90]
	d: timestamp in milliseconds since launch
	
Further information about the face tracking properties can be found [here](http://msdn.microsoft.com/en-us/library/jj130970.aspx)

### FaceAnimation message - message with the coordinates of a face event:
The messages will have the following format:

    Address pattern: "/face_animation"
    Type tag: "iiffffd"
    i: The ID of the sensor
    i: The ID of the user
    f: lip raiser [-1, 1]
    f: lip stretcher [-1, 1]
    f: lip corner depressor [-1, 1]
    f: jaw lower [-1, 1]
    f: brow lower [-1, 1]
    f: brow raiser [-1, 1]
	d: timestamp in milliseconds since launch

Further information about the AnimationUnit properties can be found [here](http://msdn.microsoft.com/en-us/library/jj130970.aspx)


### Full list of joints

* head -> head
* neck -> center shoulder
* torso -> spine
* r_collar #not supported by KinectSDK (yet)
* r_shoulder -> right shoulder
* r_elbow -> right elbow
* r_wrist -> right wrist
* r_hand -> right hand
* r_finger #not supported by KinectSDK (yet)
* l_collar #not supported by KinectSDK (yet)
* l_shoulder -> left shoulder
* l_elbow -> left elbow
* l_wrist -> left wrist
* l_hand -> left hand
* l_finger #not supported by KinectSDK (yet)
* r_hip -> right hip
* r_knee -> right knee
* r_ankle -> right ankle
* r_foot -> right foot
* l_hip -> left hip
* l_knee -> left knee
* l_ankle -> left ankle
* l_foot -> left foot


Other
-----
### For feature request, reporting bugs, or general OSCeleton 
discussion, come join the fun in a related [google group](http://groups.google.com/group/osceleton)!

### OSCeleton-OpenNI ?
To use the OpenNI & NITE framework in combination with OSC messages, download [OSCeleton-OpenNI](https://github.com/Zillode/OSCeleton-OpenNI)

### OSCeleton-KinectSDK2 ?
To use the Kinect SDK2 in combination with OSC messages, download [OSCeleton-KinectSDK2](https://github.com/Zillode/OSCeleton-KinectSDK2)

Have fun!


