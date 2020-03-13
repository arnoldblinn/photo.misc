Attribute VB_Name = "XMLIndex"
Option Explicit
Rem -------------------------------------------------------------
Rem Copyright 2000, by Arnold N. Blinn.  All rights reserved
Rem
Rem File: XMLIndex.bas
Rem
Rem Description:
Rem     Contains constants for the offset in the XML structure.
Rem
Rem -------------------------------------------------------------

Rem Task constants
Public Const XMLITask_Name = 0
Public Const XMLITask_Status = 1
Public Const XMLITask_Source = 2
Public Const XMLITask_Schedule = 3
Public Const XMLITask_Format = 4
Public Const XMLITask_Destination = 5

Public Const XMLITask_Max = 6

Rem Status constants
Public Const XMLIStatus_LastRun = 0
Public Const XMLIStatus_Failed = 1
Public Const XMLIStatus_Reason = 2

Public Const XMLIStatus_Max = 3

Rem Source constants
Public Const XMLISource_Type = 0
Public Const XMLISource_Album = 1
Public Const XMLISource_AlbumURI = 2

Public Const XMLISource_Max = 3

Public Const XMLISource_Type_Album = 0
Public Const XMLISource_Type_AlbumURI = 1

Rem Album constants
Public Const XMLIAlbum_Name = 0
Public Const XMLIAlbum_PictureList = 1

Public Const XMLIAlbum_Max = 2

Rem Picture constants
Public Const XMLIPicture_Name = 0
Public Const XMLIPicture_URI = 1

Public Const XMLIPicture_Max = 2

Rem Schedule constants
Public Const XMLISchedule_Type = 0
Public Const XMLISchedule_Disable = 1
Public Const XMLISchedule_Connect = 2
Public Const XMLISchedule_Hours = 3
Public Const XMLISchedule_Minutes = 4
Public Const XMLISchedule_Weekday = 5
Public Const XMLISchedule_Monthday = 6

Public Const XMLISchedule_Max = 7

Public Const XMLISchedule_Type_None = 0
Public Const XMLISchedule_Type_Hourly = 1
Public Const XMLISchedule_Type_Daily = 2
Public Const XMLISchedule_Type_Weekly = 3
Public Const XMLISchedule_Type_Monthly = 4

Public Const XMLISchedule_Weekday_Sunday = 0
Public Const XMLISchedule_Weekday_Monday = 1
Public Const XMLISchedule_Weekday_Tuesday = 2
Public Const XMLISchedule_Weekday_Wednesday = 3
Public Const XMLISchedule_Weekday_Thursday = 4
Public Const XMLISchedule_Weekday_Friday = 5
Public Const XMLISchedule_Weekday_Saturday = 6

Rem Format settings constants
Public Const XMLIFormatSettings_Width = 0
Public Const XMLIFormatSettings_Height = 1
Public Const XMLIFormatSettings_Grow = 2
Public Const XMLIFormatSettings_Shrink = 3
Public Const XMLIFormatSettings_Rotate = 4
Public Const XMLIFormatSettings_RotateDirection = 5
Public Const XMLIFormatSettings_Pad = 6
Public Const XMLIFormatSettings_PadColor = 7
Public Const XMLIFormatSettings_VerticalAlign = 8
Public Const XMLIFormatSettings_HorizontalAlign = 9
Public Const XMLIFormatSettings_Margins = 10
Public Const XMLIFormatSettings_TopMargin = 11
Public Const XMLIFormatSettings_LeftMargin = 12
Public Const XMLIFormatSettings_Rightmargin = 13
Public Const XMLIFormatSettings_BottomMargin = 14
Public Const XMLIFormatSettings_MarginColor = 15
Public Const XMLIFormatSettings_Compression = 16
Public Const XMLIFormatSettings_Thumbnail = 17
Public Const XMLIFormatSettings_ThumbWidth = 18
Public Const XMLIFormatSettings_ThumbHeight = 19

Public Const XMLIFormatSettings_Max = 20

Public Const XMLIFormatSettings_RotateDirection_CW = 0
Public Const XMLIFormatSettings_RotateDirection_CCW = 1

Rem Format Profile constants
Public Const XMLIFormatProfile_Name = 0
Public Const XMLIFormatProfile_Settings = 1

Public Const XMLIFormatProfile_Max = 2

Rem Format constants
Public Const XMLIFormat_Name = 0
Public Const XMLIFormat_Settings = 1

Public Const XMLIFormat_Max = 2

Rem Destination constants
Public Const XMLIDestination_Type = 0
Public Const XMLIDestination_Directory = 1
Public Const XMLIDestination_DirectoryDelete = 2
Public Const XMLIDestination_FileTemplate = 3
Public Const XMLIDestination_DigiFramePort = 4
Public Const XMLIDestination_DigiFrameMedia = 5

Public Const XMLIDestination_Max = 6

Public Const XMLIDestination_Type_Directory = 0
Public Const XMLIDestination_Type_DigiFrame = 1

Public Const XMLIDestination_DigiFramePort_COM1 = 0
Public Const XMLIDestination_DigiFramePort_COM2 = 1
Public Const XMLIDestination_DigiFramePort_COM3 = 2
Public Const XMLIDestination_DigiFramePort_COM4 = 3

Public Const XMLIDestination_DigiFrameMedia_CompactFlash = 0
Public Const XMLIDestination_DigiFrameMedia_SmartMedia = 1





