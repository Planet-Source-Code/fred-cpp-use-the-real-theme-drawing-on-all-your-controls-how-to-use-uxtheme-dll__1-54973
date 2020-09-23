VERSION 5.00
Begin VB.UserControl isThemedControl 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
End
Attribute VB_Name = "isThemedControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fred.cpp 2004
' http://mx.geocities.com/fred_cpp
'
' Description: Make your controls use the Windows Theme Style
' Warning, Themes will be visible ONLY ON WINDOWS XP,
' not In previous versions of Windows.
' Requires UxTheme.dll, Included In WinXP.

' Constants and types taken from cUxTheme.cls from
' Steve at VBAccelerator.com
Option Explicit
'' //---------------------------------------------------------------------------------------
'' //   "Window" (i.e., non-client) Parts & States
'' //
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeWindowParts
    WP_CAPTION = 1
    WP_SMALLCAPTION = 2
    WP_MINCAPTION = 3
    WP_SMALLMINCAPTION = 4
    WP_MAXCAPTION = 5
    WP_SMALLMAXCAPTION = 6
    WP_FRAMELEFT = 7
    WP_FRAMERIGHT = 8
    WP_FRAMEBOTTOM = 9
    WP_SMALLFRAMELEFT = 10
    WP_SMALLFRAMERIGHT = 11
    WP_SMALLFRAMEBOTTOM = 12
    '' //---- window frame buttons ----
    WP_SYSBUTTON = 13
    WP_MDISYSBUTTON = 14
    WP_MINBUTTON = 15
    WP_MDIMINBUTTON = 16
    WP_MAXBUTTON = 17
    WP_CLOSEBUTTON = 18
    WP_SMALLCLOSEBUTTON = 19
    WP_MDICLOSEBUTTON = 20
    WP_RESTOREBUTTON = 21
    WP_MDIRESTOREBUTTON = 22
    WP_HELPBUTTON = 23
    WP_MDIHELPBUTTON = 24
    '' //---- scrollbars
    WP_HORZSCROLL = 25
    WP_HORZTHUMB = 26
    WP_VERTSCROLL = 27
    WP_VERTTHUMB = 28
    '' //---- dialog ----
    WP_DIALOG = 29
    '' //---- hit-test templates ---
    WP_CAPTIONSIZINGTEMPLATE = 30
    WP_SMALLCAPTIONSIZINGTEMPLATE = 31
    WP_FRAMELEFTSIZINGTEMPLATE = 32
    WP_SMALLFRAMELEFTSIZINGTEMPLATE = 33
    WP_FRAMERIGHTSIZINGTEMPLATE = 34
    WP_SMALLFRAMERIGHTSIZINGTEMPLATE = 35
    WP_FRAMEBOTTOMSIZINGTEMPLATE = 36
    WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE = 37
End Enum

Public Enum UxThemeFrameStates
    FS_ACTIVE = 1
    FS_INACTIVE = 2
End Enum

Public Enum UxThemeCaptionStates
    CS_ACTIVE = 1
    CS_INACTIVE = 2
    CS_DISABLED = 3
End Enum
    
Public Enum UxThemeMaxCaptionStates
    MXCS_ACTIVE = 1
    MXCS_INACTIVE = 2
    MXCS_DISABLED = 3
End Enum

Public Enum UxThemeMinCaptionStates
    MNCS_ACTIVE = 1
    MNCS_INACTIVE = 2
    MNCS_DISABLED = 3
End Enum

Public Enum UxThemeHorzScrollStates
    HSS_NORMAL = 1
    HSS_HOT = 2
    HSS_PUSHED = 3
    HSS_DISABLED = 4
End Enum

Public Enum UxThemeHorzThumbStates
    HTS_NORMAL = 1
    HTS_HOT = 2
    HTS_PUSHED = 3
    HTS_DISABLED = 4
End Enum

Public Enum UxThemeVertScrollStates
    VSS_NORMAL = 1
    VSS_HOT = 2
    VSS_PUSHED = 3
    VSS_DISABLED = 4
End Enum

Public Enum UxThemeVertThumbStates
    VTS_NORMAL = 1
    VTS_HOT = 2
    VTS_PUSHED = 3
    VTS_DISABLED = 4
End Enum

Public Enum UxThemeSysButtonStates
    SBS_NORMAL = 1
    SBS_HOT = 2
    SBS_PUSHED = 3
    SBS_DISABLED = 4
End Enum

Public Enum UxThemeMinButtonStates
    MINBS_NORMAL = 1
    MINBS_HOT = 2
    MINBS_PUSHED = 3
    MINBS_DISABLED = 4
End Enum

Public Enum UxThemeMaxButtonStates
    MAXBS_NORMAL = 1
    MAXBS_HOT = 2
    MAXBS_PUSHED = 3
    MAXBS_DISABLED = 4
End Enum

Public Enum UxThemeRestoreButtonStates
    RBS_NORMAL = 1
    RBS_HOT = 2
    RBS_PUSHED = 3
    RBS_DISABLED = 4
End Enum

Public Enum UxThemeHelpButtonStates
    HBS_NORMAL = 1
    HBS_HOT = 2
    HBS_PUSHED = 3
    HBS_DISABLED = 4
End Enum

Public Enum UxThemeCloseButtonStates
    CBS_NORMAL = 1
    CBS_HOT = 2
    CBS_PUSHED = 3
    CBS_DISABLED = 4
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Button" Parts & States
'' //--------------------------------------------------------------------------------------
Public Enum UxThemeButtonParts
    BP_PUSHBUTTON = 1
    BP_RADIOBUTTON = 2
    BP_CHECKBOX = 3
    BP_GROUPBOX = 4
    BP_USERBUTTON = 5
End Enum

Public Enum UxThemePushButtonStates
    PBS_NORMAL = 1
    PBS_HOT = 2
    PBS_PRESSED = 3
    PBS_DISABLED = 4
    PBS_DEFAULTED = 5
End Enum

Public Enum UxThemeRadioButtonStates
    RBS_UNCHECKEDNORMAL = 1
    RBS_UNCHECKEDHOT = 2
    RBS_UNCHECKEDPRESSED = 3
    RBS_UNCHECKEDDISABLED = 4
    RBS_CHECKEDNORMAL = 5
    RBS_CHECKEDHOT = 6
    RBS_CHECKEDPRESSED = 7
    RBS_CHECKEDDISABLED = 8
End Enum

Public Enum UxThemeCheckBoxStates
    CBS_UNCHECKEDNORMAL = 1
    CBS_UNCHECKEDHOT = 2
    CBS_UNCHECKEDPRESSED = 3
    CBS_UNCHECKEDDISABLED = 4
    CBS_CHECKEDNORMAL = 5
    CBS_CHECKEDHOT = 6
    CBS_CHECKEDPRESSED = 7
    CBS_CHECKEDDISABLED = 8
    CBS_MIXEDNORMAL = 9
    CBS_MIXEDHOT = 10
    CBS_MIXEDPRESSED = 11
    CBS_MIXEDDISABLED = 12
End Enum

Public Enum UxThemeGroupBoxStates
    GBS_NORMAL = 1
    GBS_DISABLED = 2
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Rebar" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeRebarParts
    RP_GRIPPER = 1
    RP_GRIPPERVERT = 2
    RP_BAND = 3
    RP_CHEVRON = 4
    RP_CHEVRONVERT = 5
End Enum

Public Enum UxThemeChevronStates
    CHEVS_NORMAL = 1
    CHEVS_HOT = 2
    CHEVS_PRESSED = 3
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Toolbar" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeToolBarParts
    TP_BUTTON = 1
    TP_DROPDOWNBUTTON = 2
    TP_SPLITBUTTON = 3
    TP_SPLITBUTTONDROPDOWN = 4
    TP_SEPARATOR = 5
    TP_SEPARATORVERT = 6
End Enum

Public Enum UxThemeToolBarStates
    TS_NORMAL = 1
    TS_HOT = 2
    TS_PRESSED = 3
    TS_DISABLED = 4
    TS_CHECKED = 5
    TS_HOTCHECKED = 6
End Enum

'' //---------------------------------------------------------------------------------------
'' //   "Status" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeStatusParts
    SP_PANE = 1
    SP_GRIPPERPANE = 2
    SP_GRIPPER = 3
End Enum

'' //---------------------------------------------------------------------------------------
'' //   "Menu" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeMenuParts
    MP_MENUITEM = 1
    MP_MENUDROPDOWN = 2
    MP_MENUBARITEM = 3
    MP_MENUBARDROPDOWN = 4
    MP_CHEVRON = 5
    MP_SEPARATOR = 6
End Enum

Public Enum UxThemeMenuStates
    MS_NORMAL = 1
    MS_SELECTED = 2
    MS_DEMOTED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "ListView" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeLISTVIEWParts
   LVP_LISTITEM = 1
   LVP_LISTGROUP = 2
   LVP_LISTDETAIL = 3
   LVP_LISTSORTEDDETAIL = 4
   LVP_EMPTYTEXT = 5
End Enum

Public Enum UxThemeLISTITEMStates
   LIS_NORMAL = 1
   LIS_HOT = 2
   LIS_SELECTED = 3
   LIS_DISABLED = 4
   LIS_SELECTEDNOTFOCUS = 5
End Enum

' //---------------------------------------------------------------------------------------
' //   "Header" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeHEADERParts
   HP_HEADERITEM = 1
   HP_HEADERITEMLEFT = 2
   HP_HEADERITEMRIGHT = 3
   HP_HEADERSORTARROW = 4
End Enum

Public Enum UxThemeHEADERITEMStates
   HIS_NORMAL = 1
   HIS_HOT = 2
   HIS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMLEFTStates
   HILS_NORMAL = 1
   HILS_HOT = 2
   HILS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMRIGHTStates
   HIRS_NORMAL = 1
   HIRS_HOT = 2
   HIRS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERSORTARROWStates
   HSAS_SORTEDUP = 1
   HSAS_SORTEDDOWN = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Progress" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemePROGRESSParts
   PP_BAR = 1
   PP_BARVERT = 2
   PP_CHUNK = 3
   PP_CHUNKVERT = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tab" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UsxThemeTABParts
   TABP_TABITEM = 1
   TABP_TABITEMLEFTEDGE = 2
   TABP_TABITEMRIGHTEDGE = 3
   TABP_TABITEMBOTHEDGE = 4
   TABP_TOPTABITEM = 5
   TABP_TOPTABITEMLEFTEDGE = 6
   TABP_TOPTABITEMRIGHTEDGE = 7
   TABP_TOPTABITEMBOTHEDGE = 8
   TABP_PANE = 9
   TABP_BODY = 10
End Enum

Public Enum UxThemeTABITEMStates
   TIS_NORMAL = 1
   TIS_HOT = 2
   TIS_SELECTED = 3
   TIS_DISABLED = 4
   TIS_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMLEFTEDGEStates
   TILES_NORMAL = 1
   TILES_HOT = 2
   TILES_SELECTED = 3
   TILES_DISABLED = 4
   TILES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMRIGHTEDGEStates
   TIRES_NORMAL = 1
   TIRES_HOT = 2
   TIRES_SELECTED = 3
   TIRES_DISABLED = 4
   TIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMBOTHEDGESStates
   TIBES_NORMAL = 1
   TIBES_HOT = 2
   TIBES_SELECTED = 3
   TIBES_DISABLED = 4
   TIBES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMStates
   TTIS_NORMAL = 1
   TTIS_HOT = 2
   TTIS_SELECTED = 3
   TTIS_DISABLED = 4
   TTIS_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMLEFTEDGEStates
   TTILES_NORMAL = 1
   TTILES_HOT = 2
   TTILES_SELECTED = 3
   TTILES_DISABLED = 4
   TTILES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMRIGHTEDGEStates
   TTIRES_NORMAL = 1
   TTIRES_HOT = 2
   TTIRES_SELECTED = 3
   TTIRES_DISABLED = 4
   TTIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMBOTHEDGESStates
   TTIBES_NORMAL = 1
   TTIBES_HOT = 2
   TTIBES_SELECTED = 3
   TTIBES_DISABLED = 4
   TTIBES_FOCUSED = 5
End Enum

' //---------------------------------------------------------------------------------------
' //   "Trackbar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTRACKBARParts
   TKP_TRACK = 1
   TKP_TRACKVERT = 2
   TKP_THUMB = 3
   TKP_THUMBBOTTOM = 4
   TKP_THUMBTOP = 5
   TKP_THUMBVERT = 6
   TKP_THUMBLEFT = 7
   TKP_THUMBRIGHT = 8
   TKP_TICS = 9
   TKP_TICSVERT = 10
End Enum

Public Enum UxThemeTRACKBARStates
   TKS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKStates
   TRS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKVERTStates
   TRVS_NORMAL = 1
End Enum

Public Enum UxThemeTHUMBStates
   TUS_NORMAL = 1
   TUS_HOT = 2
   TUS_PRESSED = 3
   TUS_FOCUSED = 4
   TUS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBBOTTOMStates
   TUBS_NORMAL = 1
   TUBS_HOT = 2
   TUBS_PRESSED = 3
   TUBS_FOCUSED = 4
   TUBS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBTOPStates
   TUTS_NORMAL = 1
   TUTS_HOT = 2
   TUTS_PRESSED = 3
   TUTS_FOCUSED = 4
   TUTS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBVERTStates
   TUVS_NORMAL = 1
   TUVS_HOT = 2
   TUVS_PRESSED = 3
   TUVS_FOCUSED = 4
   TUVS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBLEFTStates
   TUVLS_NORMAL = 1
   TUVLS_HOT = 2
   TUVLS_PRESSED = 3
   TUVLS_FOCUSED = 4
   TUVLS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBRIGHTStates
   TUVRS_NORMAL = 1
   TUVRS_HOT = 2
   TUVRS_PRESSED = 3
   TUVRS_FOCUSED = 4
   TUVRS_DISABLED = 5
End Enum

Public Enum UxThemeTICSStates
   TSS_NORMAL = 1
End Enum

Public Enum UxThemeTICSVERTStates
   TSVS_NORMAL = 1
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tooltips" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTOOLTIPParts
   TTP_STANDARD = 1
   TTP_STANDARDTITLE = 2
   TTP_BALLOON = 3
   TTP_BALLOONTITLE = 4
   TTP_CLOSE = 5
End Enum

Public Enum UxThemeCLOSEStates
   TTCS_NORMAL = 1
   TTCS_HOT = 2
   TTCS_PRESSED = 3
End Enum

Public Enum UxThemeSTANDARDStates
   TTSS_NORMAL = 1
   TTSS_LINK = 2
End Enum

Public Enum UxThemeBALLOONStates
   TTBS_NORMAL = 1
   TTBS_LINK = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "TreeView" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTREEVIEWParts
   TVP_TREEITEM = 1
   TVP_GLYPH = 2
   TVP_BRANCH = 3
End Enum

Public Enum UxThemeTREEITEMStates
   TREIS_NORMAL = 1
   TREIS_HOT = 2
   TREIS_SELECTED = 3
   TREIS_DISABLED = 4
   TREIS_SELECTEDNOTFOCUS = 5
End Enum

Public Enum UxThemeGLYPHStates
   GLPS_CLOSED = 1
   GLPS_OPENED = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Spin" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSPINStates
   SPNP_UP = 1
   SPNP_DOWN = 2
   SPNP_UPHORZ = 3
   SPNP_DOWNHORZ = 4
End Enum

Public Enum UxThemeUPStates
   UPS_NORMAL = 1
   UPS_HOT = 2
   UPS_PRESSED = 3
   UPS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNStates
   DNS_NORMAL = 1
   DNS_HOT = 2
   DNS_PRESSED = 3
   DNS_DISABLED = 4
End Enum

Public Enum UxThemeUPHORZStates
   UPHZS_NORMAL = 1
   UPHZS_HOT = 2
   UPHZS_PRESSED = 3
   UPHZS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNHORZStates
   DNHZS_NORMAL = 1
   DNHZS_HOT = 2
   DNHZS_PRESSED = 3
   DNHZS_DISABLED = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Page" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemePageParts
   PGRP_UP = 1
   PGRP_DOWN = 2
   PGRP_UPHORZ = 3
   PGRP_DOWNHORZ = 4
End Enum

' //--- Pager uses same states as Spin ---

' //---------------------------------------------------------------------------------------
' //   "Scrollbar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSCROLLBARParts
   SBP_ARROWBTN = 1
   SBP_THUMBBTNHORZ = 2
   SBP_THUMBBTNVERT = 3
   SBP_LOWERTRACKHORZ = 4
   SBP_UPPERTRACKHORZ = 5
   SBP_LOWERTRACKVERT = 6
   SBP_UPPERTRACKVERT = 7
   SBP_GRIPPERHORZ = 8
   SBP_GRIPPERVERT = 9
   SBP_SIZEBOX = 10
End Enum



Public Enum UxThemeARROWBTNStates
   ABS_UPNORMAL = 1
   ABS_UPHOT = 2
   ABS_UPPRESSED = 3
   ABS_UPDISABLED = 4
   ABS_DOWNNORMAL = 5
   ABS_DOWNHOT = 6
   ABS_DOWNPRESSED = 7
   ABS_DOWNDISABLED = 8
   ABS_LEFTNORMAL = 9
   ABS_LEFTHOT = 10
   ABS_LEFTPRESSED = 11
   ABS_LEFTDISABLED = 12
   ABS_RIGHTNORMAL = 13
   ABS_RIGHTHOT = 14
   ABS_RIGHTPRESSED = 15
   ABS_RIGHTDISABLED = 16
End Enum

Public Enum UxThemeSCROLLBARStates
   SCRBS_NORMAL = 1
   SCRBS_HOT = 2
   SCRBS_PRESSED = 3
   SCRBS_DISABLED = 4
End Enum

Public Enum UxThemeSIZEBOXStates
   SZB_RIGHTALIGN = 1
   SZB_LEFTALIGN = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Edit" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeEDITParts
   EP_EDITTEXT = 1
   EP_CARET = 2
End Enum

Public Enum UxThemeEDITTEXTStates
   ETS_NORMAL = 1
   ETS_HOT = 2
   ETS_SELECTED = 3
   ETS_DISABLED = 4
   ETS_FOCUSED = 5
   ETS_READONLY = 6
   ETS_ASSIST = 7
End Enum

' //---------------------------------------------------------------------------------------
' //   "ComboBox" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeComboBoxParts
   CP_DROPDOWNBUTTON = 1
End Enum

Public Enum UxThemeComboBoxStates
   CBXS_NORMAL = 1
   CBXS_HOT = 2
   CBXS_PRESSED = 3
   CBXS_DISABLED = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Taskbar Clock" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeCLOCKParts
   CLP_TIME = 1
End Enum

Public Enum UxThemeCLOCKStates
   CLS_NORMAL = 1
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tray Notify" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTRAYNOTIFYParts
   TNP_BACKGROUND = 1
   TNP_ANIMBACKGROUND = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTASKBARParts
   TBP_BACKGROUNDBOTTOM = 1
   TBP_BACKGROUNDRIGHT = 2
   TBP_BACKGROUNDTOP = 3
   TBP_BACKGROUNDLEFT = 4
   TBP_SIZINGBARBOTTOM = 5
   TBP_SIZINGBARRIGHT = 6
   TBP_SIZINGBARTOP = 7
   TBP_SIZINGBARLEFT = 8
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBand" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTASKBANDParts
   TDP_GROUPCOUNT = 1
   TDP_FLASHBUTTON = 2
   TDP_FLASHBUTTONGROUPMENU = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "StartPanel" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSTARTPANELParts
   SPP_USERPANE = 1
   SPP_MOREPROGRAMS = 2
   SPP_MOREPROGRAMSARROW = 3
   SPP_PROGLIST = 4
   SPP_PROGLISTSEPARATOR = 5
   SPP_PLACESLIST = 6
   SPP_PLACESLISTSEPARATOR = 7
   SPP_LOGOFF = 8
   SPP_LOGOFFBUTTONS = 9
   SPP_USERPICTURE = 10
   SPP_PREVIEW = 11
End Enum

Public Enum UxThemeMOREPROGRAMSARROWStates
   SPS_NORMAL = 1
   SPS_HOT = 2
   SPS_PRESSED = 3
End Enum

Public Enum UxThemeLOGOFFBUTTONSStates
   SPLS_NORMAL = 1
   SPLS_HOT = 2
   SPLS_PRESSED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "ExplorerBar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeEXPLORERBARParts
   EBP_HEADERBACKGROUND = 1
   EBP_HEADERCLOSE = 2
   EBP_HEADERPIN = 3
   EBP_IEBARMENU = 4
   EBP_NORMALGROUPBACKGROUND = 5
   EBP_NORMALGROUPCOLLAPSE = 6
   EBP_NORMALGROUPEXPAND = 7
   EBP_NORMALGROUPHEAD = 8
   EBP_SPECIALGROUPBACKGROUND = 9
   EBP_SPECIALGROUPCOLLAPSE = 10
   EBP_SPECIALGROUPEXPAND = 11
   EBP_SPECIALGROUPHEAD = 12
End Enum

Public Enum UxThemeHEADERCLOSEStates
   EBHC_NORMAL = 1
   EBHC_HOT = 2
   EBHC_PRESSED = 3
End Enum

Public Enum UxThemeHEADERPINStates
   EBHP_NORMAL = 1
   EBHP_HOT = 2
   EBHP_PRESSED = 3
   EBHP_SELECTEDNORMAL = 4
   EBHP_SELECTEDHOT = 5
   EBHP_SELECTEDPRESSED = 6
End Enum

Public Enum UxThemeIEBARMENUStates
   EBM_NORMAL = 1
   EBM_HOT = 2
   EBM_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPCOLLAPSEStates
   EBNGC_NORMAL = 1
   EBNGC_HOT = 2
   EBNGC_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPEXPANDStates
   EBNGE_NORMAL = 1
   EBNGE_HOT = 2
   EBNGE_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPCOLLAPSEStates
   EBSGC_NORMAL = 1
   EBSGC_HOT = 2
   EBSGC_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPEXPANDStates
   EBSGE_NORMAL = 1
   EBSGE_HOT = 2
   EBSGE_PRESSED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBand" Parts & States
' //---------------------------------------------------------------------------------------

'Required Enums
Public Enum UxThemeMENUBANDParts
   MDP_NEWAPPBUTTON = 1
   MDP_SEPERATOR = 2
End Enum

Public Enum UxThemeMENUBANDStates
   MDS_NORMAL = 1
   MDS_HOT = 2
   MDS_PRESSED = 3
   MDS_DISABLED = 4
   MDS_CHECKED = 5
   MDS_HOTCHECKED = 6
End Enum

'Required Types

Private Type POINT
   x As Long
   Y As Long
End Type

Private Type SIZE
   cX As Long
   cY As Long
End Type

Private Type RECT
   left As Long
   tOp As Long
   Right As Long
   Bottom As Long
End Type

Private Enum THEMESIZE
    TS_MIN             '// minimum size
    TS_TRUE            '// size without stretching
    TS_DRAW            '// size that theme mgr will use to draw part
End Enum


Private Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

Private Const EDGE_SUNKEN   As Long = 10
Private Const EDGE_RAISED   As Long = 5
Private Const BF_RECT       As Long = 15

'Needed API Calls

'Open a hTheme, Needed at the begginning of the drawing
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
'Close the hTeme Handle
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
'Draw the background of the control or section
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
'Draw the parent background (for transparent and semitransparent controls with blending over the parent object
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal hdc As Long, prc As RECT) As Long
'Get the rect of the control where theme should be applyed
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECT, pContentRect As RECT) As Long
'Draw the theme text on the control.
Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
'Draw the themed Icon works With the Imagelist Object
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal hIml As Long, ByVal iImageIndex As Long) As Long
'Returns the default size of a theme data, in a THEMESIZE variable
Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, prc As RECT, ByVal eSize As THEMESIZE, psz As SIZE) As Long
'Returns the extent of the thewt drawn with the theme style
Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As DrawTextFlags, pBoundingRect As RECT, pExtentRect As RECT) As Long
'Returns true If the selected theme part is defined in the current theme
Private Declare Function IsThemePartDefined Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long) As Long

'This api call is used only If the theme is not aviable
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

'I'll use these variables to control the state of the control
Dim m_bPressed As Boolean
'And to save the control rect
Dim m_btnRect As RECT

' Description: Here is where we draw the content,
'   for a description of the values that can have
'   the parameters see the above list. this is a
'   very simple implementation of the drawing code,
'   the main porpouse Is to show how a usercontrol
'   can use the windows theme, for mopre options
'   you will need to use the other API's. you can
'   play a lot with them! Enjoy.
Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long) As Boolean
    'hTheme handle
    Dim hTheme As Long
    'Temp variable for
    Dim lResult As Long
    'If a error occurs then or we are not running XP or the visual style is windows Classic
    On Error GoTo NoXP
    'Get out hTheme Handle
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
    'Did we get a theme handle?
    If hTheme Then
        'Yes! draw the control background
        lResult = DrawThemeBackground(hTheme, UserControl.hdc, iPart, iState, m_btnRect, m_btnRect)
        'If drawing was successful, return true, or false If not.
        DrawTheme = IIf(lResult, False, True)
    Else
        'No, we couldn't get a hTheme, drawing failed
        DrawTheme = False
    End If
    'Exit the function now
    Exit Function
NoXP:
    'An Error was detected, drawing Failed
    DrawTheme = False
End Function

' Description: Set the state to pressed
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    m_bPressed = True
    UserControl_Paint
End Sub

' Description: Set the state to not pressed
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    m_bPressed = False
    UserControl_Paint
End Sub

'Update the control.
Private Sub UserControl_Paint()
    'Create a Temp var.
    Dim bDrawThemeSuccess As Boolean
    If m_bPressed Then
        'Draw the control pressed
        bDrawThemeSuccess = DrawTheme("Button", BP_PUSHBUTTON, PBS_PRESSED)
    Else
        'Draw the control Normal
        bDrawThemeSuccess = DrawTheme("Button", BP_PUSHBUTTON, PBS_NORMAL)
    End If
    'If drawing failed, then...
    If Not bDrawThemeSuccess Then
        'Use custom code to draw the control with a nonthemed look
        DrawEdge UserControl.hdc, m_btnRect, IIf(m_bPressed, EDGE_SUNKEN, EDGE_RAISED), BF_RECT
    End If
End Sub

' Description: Adjust the size of the rect
Private Sub UserControl_Resize()
    m_btnRect.left = 0
    m_btnRect.tOp = 0
    m_btnRect.Right = UserControl.ScaleWidth
    m_btnRect.Bottom = UserControl.ScaleHeight
End Sub
