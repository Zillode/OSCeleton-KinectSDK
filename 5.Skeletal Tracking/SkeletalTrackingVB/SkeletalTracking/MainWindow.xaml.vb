' (c) Copyright Microsoft Corporation.
' This source is subject to the Microsoft Public License (Ms-PL).
' Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
' All other rights reserved.

Imports System.Text
Imports Microsoft.Kinect
Imports Coding4Fun.Kinect.Wpf

Namespace SkeletalTracking
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
		End Sub

        Private _closing As Boolean = False
        Private Const skeletonCount As Integer = 6
        Private allSkeletons(skeletonCount - 1) As Skeleton

        Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            AddHandler kinectSensorChooser1.KinectSensorChanged, AddressOf kinectSensorChooser1_KinectSensorChanged

        End Sub

        Private Sub kinectSensorChooser1_KinectSensorChanged(ByVal sender As Object, ByVal e As DependencyPropertyChangedEventArgs)
            Dim old As KinectSensor = CType(e.OldValue, KinectSensor)

            StopKinect(old)

            Dim sensor As KinectSensor = CType(e.NewValue, KinectSensor)

            If sensor Is Nothing Then
                Return
            End If




            Dim parameters = New TransformSmoothParameters With {.Smoothing = 0.3F, .Correction = 0.0F, .Prediction = 0.0F, .JitterRadius = 1.0F, .MaxDeviationRadius = 0.5F}
            sensor.SkeletonStream.Enable(parameters)

            sensor.SkeletonStream.Enable()

            AddHandler sensor.AllFramesReady, AddressOf sensor_AllFramesReady
            sensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30)
            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30)

            Try
                sensor.Start()
            Catch e1 As System.IO.IOException
                kinectSensorChooser1.AppConflictOccurred()
            End Try
        End Sub

        Private Sub sensor_AllFramesReady(ByVal sender As Object, ByVal e As AllFramesReadyEventArgs)
            If _closing Then
                Return
            End If

            'Get a skeleton
            Dim first As Skeleton = GetFirstSkeleton(e)

            If first Is Nothing Then
                Return
            End If



            'set scaled position
            'ScalePosition(headImage, first.Joints[JointType.Head]);
            ScalePosition(leftEllipse, first.Joints(JointType.HandLeft))
            ScalePosition(rightEllipse, first.Joints(JointType.HandRight))

            GetCameraPoint(first, e)

        End Sub

        Private Sub GetCameraPoint(ByVal first As Skeleton, ByVal e As AllFramesReadyEventArgs)

            Using depth As DepthImageFrame = e.OpenDepthImageFrame()
                If depth Is Nothing OrElse kinectSensorChooser1.Kinect Is Nothing Then
                    Return
                End If


                'Map a joint location to a point on the depth map
                'head
                Dim headDepthPoint As DepthImagePoint = depth.MapFromSkeletonPoint(first.Joints(JointType.Head).Position)
                'left hand
                Dim leftDepthPoint As DepthImagePoint = depth.MapFromSkeletonPoint(first.Joints(JointType.HandLeft).Position)
                'right hand
                Dim rightDepthPoint As DepthImagePoint = depth.MapFromSkeletonPoint(first.Joints(JointType.HandRight).Position)


                'Map a depth point to a point on the color image
                'head
                Dim headColorPoint As ColorImagePoint = depth.MapToColorImagePoint(headDepthPoint.X, headDepthPoint.Y, ColorImageFormat.RgbResolution640x480Fps30)
                'left hand
                Dim leftColorPoint As ColorImagePoint = depth.MapToColorImagePoint(leftDepthPoint.X, leftDepthPoint.Y, ColorImageFormat.RgbResolution640x480Fps30)
                'right hand
                Dim rightColorPoint As ColorImagePoint = depth.MapToColorImagePoint(rightDepthPoint.X, rightDepthPoint.Y, ColorImageFormat.RgbResolution640x480Fps30)


                'Set location
                CameraPosition(headImage, headColorPoint)
                CameraPosition(leftEllipse, leftColorPoint)
                CameraPosition(rightEllipse, rightColorPoint)
            End Using
        End Sub


        Private Function GetFirstSkeleton(ByVal e As AllFramesReadyEventArgs) As Skeleton
            Using skeletonFrameData As SkeletonFrame = e.OpenSkeletonFrame()
                If skeletonFrameData Is Nothing Then
                    Return Nothing
                End If


                skeletonFrameData.CopySkeletonDataTo(allSkeletons)

                'get the first tracked skeleton
                Dim first As Skeleton = ( _
                    From s In allSkeletons _
                    Where s.TrackingState = SkeletonTrackingState.Tracked _
                    Select s).FirstOrDefault()

                Return first

            End Using
        End Function

        Private Sub StopKinect(ByVal sensor As KinectSensor)
            If sensor IsNot Nothing Then
                If sensor.IsRunning Then
                    'stop sensor
                    sensor.Stop()

                    'stop audio if not null
                    If sensor.AudioSource IsNot Nothing Then
                        sensor.AudioSource.Stop()
                    End If
                End If
            End If
        End Sub

        Private Sub CameraPosition(ByVal element As FrameworkElement, ByVal point As ColorImagePoint)
            'Divide by 2 for width and height so point is right in the middle 
            ' instead of in top/left corner
            Canvas.SetLeft(element, point.X - (element.Width \ 2))
            Canvas.SetTop(element, point.Y - element.Height \ 2)

        End Sub

        Private Sub ScalePosition(ByVal element As FrameworkElement, ByVal joint As Joint)
            'convert the value to X/Y
            'Joint scaledJoint = joint.ScaleTo(1280, 720); 

            'convert & scale (.3 = means 1/3 of joint distance)
            Dim scaledJoint As Joint = joint.ScaleTo(1280, 720, 0.3F, 0.3F)

            Canvas.SetLeft(element, scaledJoint.Position.X)
            Canvas.SetTop(element, scaledJoint.Position.Y)

        End Sub


        Private Sub Window_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            _closing = True
            StopKinect(kinectSensorChooser1.Kinect)
        End Sub



	End Class
End Namespace
