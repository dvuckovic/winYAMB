Attribute VB_Name = "HHelp"
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_TP_HELP_WM_HELP = &H11

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_FINDER = &H0
Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_ENUM_INFO_TYPE = &H7

Private Const HH_SET_INFO_TYPE = &H8
Private Const HH_SYNC = &H9
Private Const HH_ADD_NAV_UI = &HA
Private Const HH_ADD_BUTTON = &HB
Private Const HH_GETBROWSER_APP = &HC
Private Const HH_KEYWORD_LOOKUP = &HD
Private Const HH_DISPLAY_TEXT_POPUP = &HE

Private Const HH_TP_HELP_CONTEXTMENU = &H10
Private Const HH_CLOSE_ALL = &H12
Private Const HH_ALINK_LOOKUP = &H13
Private Const HH_GET_LAST_ERROR = &H14
Private Const HH_ENUM_CATEGORY = &H15

Private Const HH_ENUM_CATEGORY_IT = &H16
Private Const HH_RESET_IT_FILTER = &H17
Private Const HH_SET_INCLUSIVE_FILTER = &H18
Private Const HH_SET_EXCLUSIVE_FILTER = &H19
Private Const HH_SET_GUID = &H1A
Private Const HH_INTERNAL = &HFF

Private Const IDTB_EXPAND = 200
Private Const IDTB_CONTRACT = 201
Private Const IDTB_STOP = 202
Private Const IDTB_REFRESH = 203
Private Const IDTB_BACK = 204
Private Const IDTB_HOME = 205
Private Const IDTB_SYNC = 206
Private Const IDTB_PRINT = 207
Private Const IDTB_OPTIONS = 208
Private Const IDTB_FORWARD = 209
Private Const IDTB_NOTES = 210
Private Const IDTB_BROWSE_FWD = 211
Private Const IDTB_BROWSE_BACK = 212
Private Const IDTB_CONTENTS = 213
Private Const IDTB_INDEX = 214
Private Const IDTB_SEARCH = 215
Private Const IDTB_HISTORY = 216
Private Const IDTB_BOOKMARKS = 217
Private Const IDTB_JUMP1 = 218
Private Const IDTB_JUMP2 = 219
Private Const IDTB_CUSTOMIZE = 221
Private Const IDTB_ZOOM = 222
Private Const IDTB_TOC_NEXT = 223
Private Const IDTB_TOC_PREV = 224

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type tagHHN_NOTIFY
  hdr As Variant
  pszUrl As String
End Type

Private Type tagHH_POPUP
  cbStruct As Integer
  hinst As Variant
  idString As Variant
  pszText As String
  pt As Integer
  clrForeground As ColorConstants
  clrBackground As ColorConstants
  rcMargins As RECT
  pszFont As String
End Type

Private Type tagHH_AKLINK
  cbStruct As Integer
  fReserved As Boolean
  pszKeywords As String
  pszUrl As String
  pszMsgText As String
  pszMsgTitle As String
  pszWindow As String
  fIndexOnFail As Boolean
End Type

Private Enum NavigationTypes
  HHWIN_NAVTYPE_TOC
  HHWIN_NAVTYPE_INDEX
  HHWIN_NAVTYPE_SEARCH
  HHWIN_NAVTYPE_BOOKMARKS
  HHWIN_NAVTYPE_HISTORY
End Enum

Private Enum IT
  IT_INCLUSIVE
  IT_EXCLUSIVE
  IT_HIDDEN
End Enum

Private Type tagHH_ENUM_IT
  cbStruct As Integer
  iType As Integer
  pszCatName As String
  pszITName As String
  pszITDescription As String
End Type

Private Type tagHH_ENUM_CAT
  cbStruct As Integer
  pszCatName As String
  pszCatDescription As String
End Type

Private Type tagHH_SET_INFOTYPE
  cbStruct As Integer
  pszCatName As String
  pszInfoTypeName As String
End Type

Private Enum NavTabs
  HHWIN_NAVTAB_TOP
  HHWIN_NAVTAB_LEFT
  HHWIN_NAVTAB_BOTTOM
End Enum

Private Const HH_MAX_TABS = 19
Private Enum Tabs
  HH_TAB_CONTENTS
  HH_TAB_INDEX
  HH_TAB_SEARCH
  HH_TAB_BOOKMARKS
  HH_TAB_HISTORY
End Enum

Private Const HH_FTS_DEFAULT_PROXIMITY = (-1)

Private Type tagHH_FTS_QUERY
  cbStruct As Integer
  fUniCodeStrings As Boolean
  pszSearchQuery As String
  iProximity As Long
  fStemmedSearch As Boolean
  fTitleOnly As Boolean
  fExecute As Boolean
  pszWindow As String
End Type

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5

Private Type HH_WINTYPE
  cbStruct As Integer
  fUniCodeStrings As Boolean
  pszType As String
  fsValidMembers As Variant
  fsWinProperties As Variant
  pszCaption As String
  dwStyles As Variant
  dwExStyles As Variant
  rcWindowPos As RECT
  nShowState As Integer
  hwndHelp As Variant
  hwndCaller As Variant
  hwndToolBar As Variant
  hwndNavigation As Variant
  hwndHTML As Variant
  iNavWidth As Integer
  rcHTML As RECT
  pszToc As String
  pszIndex As String
  pszFile As String
  pszHome As String
  fsToolBarFlags As Variant
  fNotExpanded As Boolean
  curNavType As Integer
  tabpos As Integer
  idNotify As Integer
  tabOrder(HH_MAX_TABS + 1) As Byte
  cHistory As Integer
  pszJump1 As String
  pszJump2 As String
  pszUrlJump1 As String
  pszUrlJump2 As String
  rcMinSize As RECT
  cbInfoTypes As Integer
End Type

Private Enum Actions
  HHACT_TAB_CONTENTS
  HHACT_TAB_INDEX
  HHACT_TAB_SEARCH
  HHACT_TAB_HISTORY
  HHACT_TAB_FAVORITES
  HHACT_EXPAND
  HHACT_CONTRACT
  HHACT_BACK
  HHACT_FORWARD
  HHACT_STOP
  HHACT_REFRESH
  HHACT_HOME
  HHACT_SYNC
  HHACT_OPTIONS
  HHACT_PRINT
  HHACT_HIGHLIGHT
  HHACT_CUSTOMIZE
  HHACT_JUMP1
  HHACT_JUMP2
  HHACT_ZOOM
  HHACT_TOC_NEXT
  HHACT_TOC_PREV
  HHACT_NOTES
  HHACT_LAST_ENUM
End Enum

Private Type tagHHNTRACK
  hdr As Variant
  pszCurUrl As String
  idAction As Integer
  phhWinType As HH_WINTYPE
End Type

Public Type HH_IDPAIR
  dwControlId As Long
  dwTopicId As Long
End Type

Private Declare Function htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Sub Help(FormObj, Topic As String)
    htmlhelp FormObj.hwnd, "winYAMB.chm::/" & Topic, &H0, 0
End Sub
