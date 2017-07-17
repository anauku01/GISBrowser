using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace COMPGIS1
{
    public class FastLook
    {

        //--------------------------------------
        // Constructor
        //--------------------------------------
         public FastLook()
         {
             
         }

        //--------------------------------------
        // FastLook calls from fleng32.dll
        //--------------------------------------
        [DllImport("fleng32.dll")]
        public static extern long CloseDwgWindow(int Channel);

        public static extern long DrawDwgFile(int Channel, string FileName);

        public static extern long DrawDwgFileView(int Channel, string FileName, double XCenter, double YCenter,
                                                  double Height, double ViewpointX, double ViewpointY, double ViewpointZ);

        public static extern long DwgWindowSize(int Channel, string FileName, long Xposition, long Yposition, long Width,
                                                long Height);



        [DllImport("kernel32.dll")]
        public static extern long GetWindowsDirectoryA(string WinPath, string Size);


        //--------------------------------------
        // Initialize the FastLook Window/drawing
        //--------------------------------------
//        public void fl_initialize(Form frm, int SystVal)
//        {
//            string WinPath;
//            string FLPath1;
//            string FLUserIni;
//            long ThisVersion;
//            long StartX;
//            long StartY;
//            long Wdth;
//            long Hght;
//            long FastLookVersion = 14;
//            long Ret;

//            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Kamel Software\FastLook");
//            ThisVersion = Convert.ToInt32(registryKey.GetValue("CurVer").ToString());
//            if (ThisVersion == FastLookVersion)
//            {
//                FLPath1 = registryKey.GetValue("Path").ToString();
//            }

        
//        If FLPath$ = "" Then
//            FLUserIni$ = GetSpecialfolderPath(CSIDL_APPDATA)
//            If FLUserIni$ <> "" Then
//                FLUserIni$ = FLUserIni$ & "\Kamel Software\FastLook V14\FLUSER.INI"
//                If Dir(FLUserIni$) <> "" Then
//                    Ret = GetPrivateProfileStringA("FASTLOOK14", "fllaunch", "", FLPath1, 256, FLUserIni$)
//                    If Ret > 1 Then
//                        FLPath$ = FnPath(ConvertString(FLPath1))
//                    End If
//                End If
//            End If
//        End If
//        If FLPath$ = "" Then
//            FastLookVersion = FastLookVersion - 1
//            GoTo tryAgain
//        Else
//            FLEnginePath$ = FLPath$ & "\fleng32.dll"
//        End If
//    End If

//        }

  
//tryAgain:
//    If FastLookVersion < 12 Then ' must find FastLook from an entry in the FLUSER.INI file in the Windows directory
//         ' ** find path to Windows for INI
//          Ret = GetWindowsDirectoryA(WinPath, 256)
//          WinPath1$ = ConvertString(WinPath$)
//          FLUserIni$ = WinPath1$ & "\FLUSER.INI"
//            'MsgBox "FLUserIni$: " & FLUserIni$
        
//         ' ** find path to FastLook
//          defaultStr$ = ""
//          IniField$ = "fllaunch"
//          FieldSection$ = "FASTLOOK12"
//          Ret = GetPrivateProfileStringA(FieldSection$, IniField$, defaultStr$, FLPath1, 256, FLUserIni$)
//          FLPath$ = ConvertString(FLPath1)
//          If FLPath$ = "" Then
//              MsgBox "Using FastLook version 11.0!"
//              FieldSection$ = "FASTLOOK11"
//              Ret = GetPrivateProfileStringA(FieldSection$, IniField$, defaultStr$, FLPath1, 256, FLUserIni$)
//              FLPath$ = ConvertString(FLPath1)
//              If FLPath$ = "" Then
//                  MsgBox "Could not locate the path to FastLook!" & Chr$(10) & "Run FastLook at least once and try again."
//                  End
//              End If
//          End If
//          FLEnginePath$ = FnPath(FLPath$) & "fleng32.dll"
//            'MsgBox "FLEnginePath$: " & FLEnginePath$
//    ElseIf FastLookVersion = 12 Then
//        'FastLook V12 posts its path in the registry, when it runs
//        ThisVersion = Val(GetRegistryStringValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Kamel Software\FastLook", "CurVer"))
//        If ThisVersion > 0 Then
//            FastLookVersion = ThisVersion
//            FLPath$ = GetRegistryStringValue("HKEY_LOCAL_MACHINE", "Software\Kamel Software\FastLook", "Path")
//        Else
//            FastLookVersion = FastLookVersion - 1
//            GoTo tryAgain
//        End If
//        FLEnginePath$ = FLPath$ & "fleng32.dll"
         
        
//    ElseIf FastLookVersion = 13 Then
//        FLUserIni$ = GetSpecialfolderPath(CSIDL_LOCAL_APPDATA)

//        If FLUserIni$ <> "" Then
//            FLUserIni$ = FLUserIni$ & "\Kamel Software\FastLook 2009\FLUSER.INI"
//            If Dir(FLUserIni$) <> "" Then
//                Ret = GetPrivateProfileStringA("FASTLOOK13", "fllaunch", "", FLPath1, 256, FLUserIni$)
//                If Ret > 1 Then
//                    FLPath$ = FnPath(ConvertString(FLPath1))
//                    FLEnginePath$ = FLPath$ & "fleng32.dll"
//                End If
//            End If
//        End If
//        If FLPath$ = "" Then
//            FLUserIni$ = GetSpecialfolderPath(CSIDL_COMMON_APPDATA)
//            If FLUserIni$ <> "" Then
//                FLUserIni$ = FLUserIni$ & "\Kamel Software\FastLook 2009\FLUSER.INI"
//                If Dir(FLUserIni$) <> "" Then
//                    Ret = GetPrivateProfileStringA("FASTLOOK13", "fllaunch", "", FLPath1, 256, FLUserIni$)
//                    If Ret > 1 Then
//                        FLPath$ = FnPath(ConvertString(FLPath1))
//                        FLEnginePath$ = FLPath$ & "fleng32.dll"
//                    End If
//                End If
//            End If
//        End If
//        If FLPath$ = "" Then
//            FastLookVersion = FastLookVersion - 1
//            GoTo tryAgain
//        End If
        
//    ElseIf FastLookVersion = 14 Then
    
//        ThisVersion = Val(GetRegistryStringValue("HKEY_CURRENT_USER", "SOFTWARE\Kamel Software\FastLook", "CurVer"))
//        If ThisVersion = FastLookVersion Then
//            FLPath$ = GetRegistryStringValue("HKEY_CURRENT_USER", "Software\Kamel Software\FastLook", "Path")
//        End If
        
//        If FLPath$ = "" Then
//            FLUserIni$ = GetSpecialfolderPath(CSIDL_APPDATA)
//            If FLUserIni$ <> "" Then
//                FLUserIni$ = FLUserIni$ & "\Kamel Software\FastLook V14\FLUSER.INI"
//                If Dir(FLUserIni$) <> "" Then
//                    Ret = GetPrivateProfileStringA("FASTLOOK14", "fllaunch", "", FLPath1, 256, FLUserIni$)
//                    If Ret > 1 Then
//                        FLPath$ = FnPath(ConvertString(FLPath1))
//                    End If
//                End If
//            End If
//        End If
//        If FLPath$ = "" Then
//            FastLookVersion = FastLookVersion - 1
//            GoTo tryAgain
//        Else
//            FLEnginePath$ = FLPath$ & "\fleng32.dll"
//        End If
//    End If
        
    
//  If Dir(FLEnginePath$) = "" Then 'can't find FastLook
//        Msg = "FastLook must have been deleted from the last location!" & Chr(10)
//        Msg = Msg & "  Please successfully run FastLook and then try again."
//        MsgBox Msg
//        Exit Function
//  End If
//  FLEngineHndl = LoadLibraryA(FLEnginePath$)
  
//' ********************************************************************

//' *** Set the OnOpen Event for your form = this procedure. "=startfl()"
//' This procedure establishes a FastLook window in the Access form.
//' Any reference to "flform1" in this example would have to be changed to your
//'  form name if you copy this code into your form.

//' ***********************************************************************

//Channel = 60  ' this sets a channel number to be referenced in any future
//              ' calls to this FastLook window.

//'*************************************************************************
//'This section is to poll the video driver for the Pixels-per-inch it displays.
//'FastLook uses pixels for window size and mouse responses and Access uses Twips.
//' We must therefore develope a Pixel per Twip conversion in order to relate
//'   our window dimensions to Access objects and to interpret our mouse events
//'   in pixels.



//Screen.MousePointer = 1

//hDC = GetDC(Frm.hWnd)

//Xpixperin = GetDeviceCaps(hDC, 88)
//Ypixperin = GetDeviceCaps(hDC, 90)
//Dummy = ReleaseDC(hDC, Frm.hWnd)

//XpixperTwip = Xpixperin / 1440
//YpixperTwip = Ypixperin / 1440
//'************************************************************

//' This section is to find the locations of the buttons in our example that
//'  border the upper left and lower right of the FastLook window in our
//'  example.  We then use these locations with the Pixel-per-Twip conversion
//'  shown earlier to set the starting point and size of the FastLook window in pixels
//'  such that it will border these objects.  In this example, as you relocate these
//'  bordering buttons, the FastLook window will follow.  If you simply desire a
//'  fixed location and size, you may skip to the OpenDwgWindow call below and put
//'  specific numbers in place of the X%,Y%,X1%,Y1% variables.  To establish a window
//'  that starts in the upper left corner of you form and is 100 pixels high and 100
//'  pixels wide, change the call to read:
//'
//'       Success = OpenDwgWindow(Forms!flform1.hWnd, Channel, 0, 0, 100, 100)
//'
//'   "flform1" should be replaced with your form name if you have placed this code
//'   in your form.

//   top1 = Frm.boxFrame1.Top
//   left1 = Frm.boxFrame1.Left
//   Height1 = Frm.boxFrame1.Height

//   top2 = Frm.boxFrame2.Top
//   left2 = Frm.boxFrame2.Left
//   width2 = Frm.boxFrame2.Width

//'   top1 = Forms!flform1.Button0.Top
//'   left1 = Forms!flform1.Button0.Left
//'   height1 = Forms!flform1.Button0.Height
  
//'   top2 = Forms!flform1.Button15.Top
//'   left2 = Forms!flform1.Button15.Left
//'   width2 = Forms!flform1.Button15.Width


//dX# = left1
//dY# = top1 + Height1
//dX# = dX# * XpixperTwip
//dY# = dY# * YpixperTwip


//StartX = dX# + 1
//StartY = dY# + 1

//dx1# = left2 + width2
//dy1# = top2
//dx1# = dx1# * XpixperTwip
//dy1# = dy1# * YpixperTwip

//Wdth = dx1# - StartX
//Hght = (dy1# - StartY) - 1
//' ********************************************************************
//success = OpenDwgWindow(Frm.hWnd, Channel, StartX, StartY, Wdth, Hght)

//    ' If you have been bounced to this location when trying to run
//    '  this example, Access was unable to locate the FLENGINE.DLL.
//    '  Check the declaration section of this module for the correct
//    '  path to FLENGINE.DLL or place FastLook in your Path Statement.
 
//' **********************************************************************
//' This section is to ask FastLook if redlining or linking are on!  Many other
//'   status queries are available - see your Help menu for this call.

//  redline = GetFLStatus(Channel, 1)
//  link = GetFLStatus(Channel, 2)

//' ***********************************************************************
//' This section is to specifically set the reponse FastLook will send to
//' your application in the form of fake mouse clicks when certain events
//' happen in the FastLook window.

//success = SetClickSettings(Channel, 0, 0, 0)  'This turns all fake mouse events off,
//                                              '  to allow you to receive only the
//                                              '  ones you want. You may elect to
//                                              '  skip all the SetClickSettings calls
//                                              '  and use the defaults given in your
//                                              '  Help menu.
                                              
//success = SetClickSettings(Channel, 2, -2, -2)
//success = SetClickSettings(Channel, 3, -3, -3)
//success = SetClickSettings(Channel, 5, -4, -4)
//success = SetClickSettings(Channel, 10, -10, -10)


//'*******************************************************************
//' Misc settings
//' Remove the scrollbars
//success = SetFLConfig(Channel, 37, 0, "")


//' If you wish to have FastLook immediately view a file, issue the following
//' call.  Fill in the variable Image$ with the full path, filename and extension
//' for the image you wish to view.  You might do this by seting this variable equal
//' to a field or combination of fields from your data record.

//'                   Success = DrawDwgFile(Channel, Image$)
//'*******************************************************************
//'START - ADDED BY MARK
//' Set the CAD Font Path - where the viewer will look for any required fonts/shapes
//' referenced in an AutoCAD drawing
//' FontPath=xxxxxx in FLUser.ini
//'*******************************************************************
//If Len(CADFontPath) > 0 Then
//    If SetFLConfig(Channel, 52, 1, CADFontPath) <> 0 Then
//    End If
//End If

//' Set the CAD External Reference Path - where the viewer will look for any required references
//' referenced in an AutoCAD drawing
//' ExRefPath=xxxxxx in FLUser.ini
//'*******************************************************************
//If Len(CADExtRefPath) > 0 Then
//    If SetFLConfig(Channel, 5, 1, CADExtRefPath) <> 0 Then
//    End If
//End If
//'END - ADDED BY MARK
//fl_initialize_Exit:
//    Exit Function

//fl_initialize_Error:
//    If Err.Number = 6 Then
//       Resume Next
//    End If
//    MsgBox Err.Number & vbCr & Err.Description
//    Resume fl_initialize_Exit

//End Function




















    }
}