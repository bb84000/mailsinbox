{*******************************************************************************
 MailInBox main unit
 bb - sdtp - april 2025
 Check mails on pop3 and imap servers
 15/04/2025 : Use array of icons to animate tray icon
              use lazbb component for FPTimer, chkmailTimer, timeTimer
 18/04/2025 : Reverse to imagelist for animated tray icon during check mail
 19/04/2025 : Revised select item event in account list
 *******************************************************************************}

unit mailsinbox1;

{$mode objfpc}{$H+}

interface

uses
  {$IFDEF WINDOWS}
  Win32Proc,
  {$ENDIF} LMessages, Classes, SysUtils, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls, Grids, ComCtrls, Buttons, Menus, IdPOP3, IdSSLOpenSSL,
  LCLIntf, IdExplicitTLSClientServerBase, IdMessage, IdIMAP4, accounts1,
  lazbbutils, FileUtil, lazbbinifiles, LazUTF8, settings1, lazbbautostart,
  lazbbaboutdlg, lazbbUpdateDlg, Impex1, mailclients1, uxtheme, Types,
  IdComponent, fptimer, RichMemo, variants, IdMessageCollection, UniqueInstance,
  log1, translations, lazbbOsVersion, lazbbcontrols, registry, dateutils,
  strutils, fpopenssl, openssl, opensslsockets;

const
  // Message post at the end of activation procedure, processed once the form is shown
  WM_FORMSHOWN = WM_USER + 1;

type
  TSaveMode = (None, Setting, All);           // Save nothing, only settings, settings and accounts
  TBtnSize=(bzLarge, bzSmall, bzNone);        // Large buttons, Small buttons, boting
  TCloseMode = (cmHide, cmQuit);              // Hide window, close application
  TFireMode = (fmLast, fmNext);               // last check, next chack
  TBtnProps = record                          // Element of buttoons array to remember its state and glyph
    Btn: TSpeedButton;
    Bmp: TBitmap;
    Enabled: Boolean;
  end;

  { TFMailsInBox }
  TFMailsInBox = class(TForm)
    ILChkMail: TImageList;
    Label1: TLabel;
    clickTimer: TLFPTimer;        // Timer used to differentiate button's single and double click
    ChkMailTimer: TLFPTimer;
    TimeTimer: TLFPTimer;
    OsVersion: TbbOsVersion;
    BtnAbout: TSpeedButton;
    BtnHelp: TSpeedButton;
    BtnAddAcc: TSpeedButton;
    BtnDeleteAcc: TSpeedButton;
    BtnEditAcc: TSpeedButton;
    BtnGetAccMail: TSpeedButton;
    BtnGetAllMail: TSpeedButton;
    BtnLaunchClient: TSpeedButton;
    BtnAccountLog: TSpeedButton;
    BtnLog: TSpeedButton;
    BtnQuit: TSpeedButton;
    BtnSettings: TSpeedButton;
    GBInfos: TGroupBox;
    IdIMAP4_1: TIdIMAP4;
    IdPOP3_1: TIdPOP3;
    ILTray: TImageList;
    ILMnuTray: TImageList;
    ILMnuMails: TImageList;
    ILMnuAccounts: TImageList;
    ILLargeBtns: TImageList;
    ILSmallBtns: TImageList;
    ILSGTitles: TImageList;
    ILMainMnu: TImageList;
    ImgAccounts: TImageList;
    LNow: TLabel;
    LStatus: TLabel;
    LVAccounts: TListView;
    MainMnu: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MMnuLog: TMenuItem;
    MMnuAbout: TMenuItem;
    MMnuHelp: TMenuItem;
    MMnuEditAcc: TMenuItem;
    MMnuAddAcc: TMenuItem;
    MMnuSettings: TMenuItem;
    MMnuGetAccMails: TMenuItem;
    MMnuGetAllMails: TMenuItem;
    MMnuPrefs: TMenuItem;
    MMnuInfos: TMenuItem;
    MMnuMails: TMenuItem;
    MMnuDisplayBar: TMenuItem;
    MMnuDisplayMenu: TMenuItem;
    MMnuDisplay: TMenuItem;
    MnuDisplayMenu: TMenuItem;
    MnuDisplayBar: TMenuItem;
    MMmnuQuit: TMenuItem;
    MMnuLaunchClient: TMenuItem;
    MMnuFile: TMenuItem;
    MMnuImport: TMenuItem;
    MnuLaunchClient: TMenuItem;
    MnuDeleteAcc: TMenuItem;
    MnuEditAcc: TMenuItem;
    MnuGetAccMail: TMenuItem;
    MnuAccountLog: TMenuItem;
    N1: TMenuItem;
    MnuAnswerMsg: TMenuItem;
    MnuDeleteMsg: TMenuItem;
    MnuGetAllMail: TMenuItem;
    MnuIconize: TMenuItem;
    MenuItem2: TMenuItem;
    MnuAbout: TMenuItem;
    MnuMaximize: TMenuItem;
    MnuMoveUp: TMenuItem;
    MnuQuit: TMenuItem;
    MnuRestore: TMenuItem;
    MnuInfos: TMenuItem;
    MnuMoveDown: TMenuItem;
    PnlAccounts: TPanel;
    PnlInfos: TPanel;
    PnlMails: TPanel;
    PnlLeft: TPanel;
    PnlStatus: TPanel;
    PnlToolbar: TPanel;
    BtnImport: TSpeedButton;
    MnuAccount: TPopupMenu;
    MnuMails: TPopupMenu;
    MnuTray: TPopupMenu;
    MnuButtonBar: TPopupMenu;
    RMInfos: TRichMemo;
    SGMails: TStringGrid;
    SplitterV: TSplitter;
    SplitterH: TSplitter;
    GetMailTimer: TTimer;
    TrayTimer: TTimer;
    TrayMail: TTrayIcon;
    UniqueInstance1: TUniqueInstance;
    procedure BtnAboutClick(Sender: TObject);
    procedure BtnDeleteAccClick(Sender: TObject);
    procedure BtnGetAccMailClick(Sender: TObject);
    procedure BtnGetAllMailClick(Sender: TObject);
    procedure BtnHelpClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure BtnLaunchClientClick(Sender: TObject);
    procedure BtnLogClick(Sender: TObject);
    procedure BtnQuitClick(Sender: TObject);
    procedure BtnQuitDblClick(Sender: TObject);
    procedure BtnQuitMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BtnSettingsClick(Sender: TObject);
    procedure ChkMailTimerStartTimer(Sender: TObject);
    procedure OnclickTimer(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure DoChangeBounds(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnEditAccClick(Sender: TObject);
    procedure BtnCloseClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormWindowStateChange(Sender: TObject);
    procedure GetMailTimerTimer(Sender: TObject);
    procedure Id_client_Connected(Sender: TObject);
    procedure Id_client_Disconnected(Sender: TObject);
    procedure IdPOP3_1Status(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure LVAccountsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure MMnuClick(Sender: TObject);
    procedure MnuAccountPopup(Sender: TObject);
    procedure MnuAnswerMsgClick(Sender: TObject);
    procedure MnuButtonBarPopup(Sender: TObject);
    procedure MnuDeleteMsgClick(Sender: TObject);
    procedure MnuDisplayBarClick(Sender: TObject);
    procedure MnuDisplayMenuClick(Sender: TObject);
    procedure MnuInfosClick(Sender: TObject);
    procedure MnuMailsPopup(Sender: TObject);
    procedure MnuMaximizeClick(Sender: TObject);
    procedure MnuIconizeClick(Sender: TObject);
    procedure MnuMoveClick(Sender: TObject);
    procedure MnuQuitClick(Sender: TObject);
    procedure MnuRestoreClick(Sender: TObject);
    procedure MnuTrayPopup(Sender: TObject);
    procedure SGMailsBeforeSelection(Sender: TObject; aCol, aRow: Integer);
    procedure SGMailsClick(Sender: TObject);
    procedure SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
      aRect: TRect; aState: TGridDrawState);
    procedure SGMailsEnter(Sender: TObject);
    procedure SGMailsExit(Sender: TObject);
    procedure SGMailsKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SGMailsMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure OnTimeTimer(Sender: TObject);
    procedure OnTrayTimer(Sender: TObject);
    procedure OnChkMailTimer(Sender: TObject);
  private
    Initialized: boolean;
    OS, OSTarget: string;
    CompileDateTime: TDateTime;
    MIBAppDataPath: string;
    MIBExecPath: string;
    ProgName: string;
    CurLangStr: string;
    LangFile: TBbIniFile;
    LangNums: TStringList;
    LangFound: boolean;
    SettingsChanged: boolean;
    AccountsChanged:Boolean;
    ConfigFileName, AccountsFileName: string;
    ChkMailTimerTick: integer;
    CanClose: boolean;
    version: string;
    sEmailCaption, sLastCheckCaption, sNextCheckCaption : string;
    sNoLongerChkUpdates, sUpdateAlertBox: string;
    OKBtn, YesBtn, NoBtn, CancelBtn: string;
    sBtnLogHint, sBtnGetAccMailHint, sBtnDeleteHint, sBtnEditAccHint: string;
    BtnsArr : array of TBtnProps;
    sAccountImported, sAccountsImported: string;
    CheckingMail: boolean;
    SGHasFocus: boolean;
    sMsgFound, sMsgsFound: string;
    TrayTimerTick: integer;
    TrayTimerBmp: TBitmap;
    sLStatusCaption: String;
    MainLog: String;
    SessionLog: TStringList;
    CurAccPend: integer;
    sCheckingAccMail, sCheckingAllMail: string;
    sConnectToServer, sDisconnectServer: string;
    sConnectedToServer, sDisconnectedServer: string;
    sMsgDeleted, sMsgNotDeleted: string;
    sConnectErrorMsg, sHeaderErrorMsg: string;
    sAccountChanged, sAccountAdded: string;
    sAccountDisabled, sAccountEnabled, AccountStatus: string;
    LogFileName: string;
    sLoadingAccounts: string;
    Iconized: Boolean;
    PrevTop, PrevLeft: integer;
    sTrayHintNoMsg, sTrayHintMsg, sTrayHintMsgs: string;
    sTrayHintNewMsg, sTrayHintNewMsgs: string;
    sTrayBallHintMsg, sTrayBallHintMsgs: string;
    sAlertBoxCBNoShowAlert: string;
    sNoCloseAlert, sNoQuitAlert, sNoShowCloseAlert, sNoShowQuitAlert: string;
    sColumnswidth: string;
    sOpenProgram, sRestart: string;
    sRetConfBack, sCreNewConf, sLoadConf: string;
    sCannotQuit, sClosingProg: string;
    sSettingsChange: string;
    StatusFmtSets: TFormatSettings;
    sMnuDelMsg, sMnuAnswerMsg: string;
    sAlertDelMmsg: string;
    sBytes, sKBytes, SMBytes: string;
    AccountPictures, TrayPicture, MailPictures, LaunchPicture: TPicture;
    slLastFires, slNextFires: TStringList;
    sNeverChecked: string;
    sUse64bit: string;
    sBtnLaunchClientDef, sBtnLaunchClientCust: string;
    sDeleteAccount: string;
    sPlsSelectAcc, sToDisplayLog, sToDeleteAcc, sToEditAcc: string;
    sAccDeleted: String;
    DisplayMails: TMailsList;
    HttpErrMsgNames: array [0..16] of string;
    sCannotGetNewVerList: string;
    sNewAccount: string;
    aMailsList: array of string;
    sCreatedDataFolder: String;
    ChkMailImgCount: Integer;

    doubleclick: Boolean;
    timespan: int64;
    lastclick: int64;
    doubleClickMaxTime: int64;
    ChkVerInterval: Int64;
    StartMini: Boolean;
    procedure Initialize;
    procedure LoadSettings(Filename: string);
    procedure SettingsOnChange(Sender: TObject);
    procedure SettingsOnStateChange(Sender: TObject);
    procedure AccountsOnChange(Sender: TObject);
    function SaveConfig(Typ: TSaveMode): boolean;
    procedure PopulateAccountsList(notify: boolean);
    procedure Translate(LngFile: TBbInifile);
   // procedure ModLangue;
    procedure SetSmallBtns(small: TBtnSize);
    function GetPendingMail(index: integer): integer;
    procedure SetProtocolProperties(CurAcc: TAccount);
    function ConnectServer(CurAcc: TAccount; var ErrorsStr: string): Integer;
    function GetHeader(CurAcc: TAccount; MailIndex: integer; var mails: TMailsList; var ErrorsStr: string): boolean;
    procedure PopulateMailsList(index: integer);
    procedure EnableControls(Enable: boolean);
    procedure UpdateInfos;
    procedure DrawTheIcon(Bmp: TBitmap; NewCount: integer; CircleColor: TColor);
    function MailChecking(status: boolean): boolean;
    procedure GetMailInfos(CurAcc: TAccount; var Mail: TMail; IdMsg: TIdMessage; siz: Integer);
    function LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
    procedure GetAccMail(ndx: integer);
    function HideOnTaskbar: boolean;
    procedure OnAppMinimize(Sender: TObject);
    procedure OnQueryendSession(var Cancel: Boolean);
    procedure GetMnuImage(ilOut: TImageList; value: variant; twin: boolean=true; ilIn: TImageList=nil);
    function GetFire(CurAcc: Taccount; mode: TFireMode): TDateTime;
    procedure SetFire(Curacc: TAccount; datim: TDateTime; mode: TFireMode);
    procedure InitButtons;
    procedure CheckUpdate(days: PtrInt);
    procedure SortMails(CurCol: integer);
    procedure LoadAccounts(filename: string);
    function SetError(E: Exception; ErrorStr: String; ErrorUID: Integer; ErrorCaption: String; var ErrorsStr: String): boolean;
    procedure BeforeClose;
    procedure OnFormShown(var Msg: TLMessage); message WM_FORMSHOWN;
    procedure InitAboutBox;
  public

    UserAppsDataPath: string;  //used by Impex1
  end;

  // Button glyphs constants
  // to better document code. Need change if add buttons and glyphs
  // Use to be sure buttons and glyphs match
Const
  glImport= 0;
  glAccountLog= 1;
  glGetAllMail= 2;
  glGetAccMail= 3;
  glLaunchClient= 4;
  glDeleteAccount= 5;
  glAddAccount= 6;
  glEditAccount= 7;
  glSettings= 8;
  glLog= 9;
  glHelp= 10;
  glAbout= 11;
  glQuit= 12;
  glLaunchOutlook= 13;
  glLaunchTbird= 14;
  glLaunchGmail= 15;
  glLaunchWin10mail= 16;
  glLaunchOutlkcom= 17;
  glLaunchCustom=18; // Future version, can choose image

var
  FMailsInBox: TFMailsInBox;

implementation

{$R *.lfm}

// Procedure to answer post message at the end of activation procedure
// so once form is shown

procedure TFMailsInBox.OnFormShown(var Msg: TLMessage);
begin
  if StartMini then Application.minimize;
  StartMini:= false;
end;

// Intercept minimize system system command to correct
// wrong window placement on restore from tray

procedure TFMailsInBox.OnAppMinimize(Sender: TObject);
begin
  if FSettings.Settings.HideInTaskbar then
  begin
    PrevLeft:=self.left;
    PrevTop:= self.top;
    WindowState:= wsMinimized;
    Iconized:= HideOnTaskbar;
  end;
end;

// Intercept end session : save all pending settings and load the program
// on next logon (Windows only for the moment. On linux, we use autostart
// and we delete auttostart on next startup

procedure TFMailsInBox.OnQueryendSession(var Cancel: Boolean);
var
  {$IFDEF WINDOWS}
    reg: TRegistry;
    RunRegKeyVal, RunRegKeySz: string;
  {$ENDIF}
  begin
  if not FSettings.Settings.Startup then
  begin
    FSettings.Settings.Restart:= true;
    {$IFDEF WINDOWS}
      reg := TRegistry.Create;
      reg.RootKey := HKEY_CURRENT_USER;
      reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\RunOnce', True) ;
      RunRegKeyVal:= UTF8ToAnsi(ProgName);
      RunRegKeySz:= UTF8ToAnsi('"'+Application.ExeName+'"');
      reg.WriteString(RunRegKeyVal, RunRegKeySz) ;
      reg.CloseKey;
      reg.free;
    {$ENDIF}
    {$IFDEF Linux}
       SetAutostart(ProgName, Application.exename);
    {$ENDIF}
  end;
  BeforeClose;
  Application.ProcessMessages;
end;

// TFMailsInBox : This is the main form of the program

procedure TFMailsInBox.FormCreate(Sender: TObject);
var
  s: string;
//  i: Integer;
 {$IFDEF Linux}
    x: Integer;
 {$ENDIF}
 tlid: TLanguageID;
begin
  // Variables initialization
  CanClose:= false;
  // Flag needed to execute once some processes in Form activation
  Initialized := False;
  MainLog:= '';
  CompileDateTime:= StringToTimeDate({$I %DATE%}+' '+{$I %TIME%}, 'yyyy/mm/dd hh:nn:ss');
  OS := 'Unk';
  ProgName := 'MailsInBox';
  CheckingMail:= false;     // Mail checking process flag
  iconized:= false;
  // DateTime settings for status bar
  StatusFmtSets:= DefaultFormatSettings ;
  StatusFmtSets.ShortDateFormat:= DefaultFormatSettings.LongDateFormat;
  // Intercept minimize system command
  Application.OnMinimize:= @OnAppMinimize;
  Application.OnQueryEndSession:= @OnQueryendSession;
  // Click timer to differentiate button single and double click
  clickTimer.Interval:= GetDoubleClickTime;
  // set clickTimer.UseTimerThread:= false;   // important !!!
  doubleClickMaxTime:= GetDoubleClickTime;
  // Initialize check mail timer
  // set ChkMailTimer.UseTimerThread:= true;
  ChkMailTimerTick:= 0;
  // Initialize time display timer
  // set TimeTimer.UseTimerThread:= true;
  TimeTimer.StartTimer;
  TrayTimerTick:=0;
  TrayTimerBmp:= TBitmap.Create;
  LangNums := TStringList.Create;
  SessionLog:= TStringList.Create;
  slLastFires:= TstringList.Create;
  slNextFires:=TstringList.Create;
  DisplayMails:= TMailsList.Create;
  // Some useful paths
  MIBExecPath:=ExtractFilePath(Application.ExeName);
  // Chargement du fichier de langue...
  LangFile:= TBbIniFile.Create(ExtractFilePath(Application.ExeName) + 'lang'+PathDelim+'fr.lng');
  {$IFDEF CPU32}
     OSTarget := '32 bits';
  {$ENDIF}
  {$IFDEF CPU64}
     OSTarget := '64 bits';
  {$ENDIF}
  {$IFDEF Linux}
    OS := 'Linux';
    CurLangStr := GetEnvironmentVariable('LANG');
    x := pos('.', CurLangStr);
    CurLangStr := Copy(CurLangStr, 0, 2);
    wxbitsrun := 0;
    //OSTarget:= '';
    UserAppsDataPath := GetUserDir;
    // Get mail client
  {$ENDIF}
  {$IFDEF WINDOWS}
    OS := 'Windows ';
    // get user data folder
    s := ExtractFilePath(ExcludeTrailingPathDelimiter(GetAppConfigDir(False)));
    if Ord(WindowsVersion) < 7 then
      UserAppsDataPath := s                     // NT to XP
    else
    UserAppsDataPath := ExtractFilePath(ExcludeTrailingPathDelimiter(s)) + 'Roaming'; // Vista to W10
    //LazGetShortLanguageID(CurLangStr);     //Deprecated
    tlid:= GetLanguageID;
    CurLangStr:= tlid.LanguageCode;
  {$ENDIF}
  version := GetVersionInfo.ProductVersion;
  // Cannot call Translate as components are not yet created, use default language
  sOpenProgram:=LangFile.ReadString('main','OpenProgram','Ouverture de Courrier en attente');
  LogAddLine(-1, now, sOpenProgram+' - Version '+Version+ ' (' + OS + OSTarget + ')');
  LogAddLine(-1, now, OSVersion.VerDetail);
  MIBAppDataPath := UserAppsDataPath + PathDelim + ProgName + PathDelim;
  if not DirectoryExists(MIBAppDataPath) then
  begin
    sCreatedDataFolder:=LangFile.ReadString(CurLangStr,'CreatedDataFolder','Dossier de données de Couriels en attente "%s" créé');
    CreateDir(MIBAppDataPath);
    LogAddLine(-1, now, Format(sCreatedDataFolder, [MIBAppDataPath]));
  end;
  LogFileName:= MIBAppDataPath+ProgName+'.log';
end;

// Add a line to log file. quote char is '|' to preserve real double quotes in log
// when reading with csv reader

function TFMailsInBox.LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
begin
  // values separated by '|'
  try
    result:= SessionLog.Add(Inttostr(acc)+'|'+TimeDateToString(dat)+'|'+evnt);
  except
    result:= SessionLog.Add('Exception error');
  end;
end;

// Form activation only needed once

procedure TFMailsInBox.FormActivate(Sender: TObject);
begin
  if not Initialized then
  begin
    InitButtons;
    Initialize;
    Application.ProcessMessages;
    if StartMini then PostMessage(Handle, WM_FORMSHOWN, 0, 0) ;
    Application.QueueAsyncCall(@CheckUpdate, ChkVerInterval);       // async call to let icons loading
  end;
end;

// populate menu imagelist from resource name(value is a string)
// or from buttons imagelist of double images (value is integer)
// Twin= true : image for enabled and disabled item

procedure TFMailsInBox.GetMnuImage(ilOut: TImageList; value: variant; twin: boolean=true; ilIn: TImageList=nil);
var
  Pict: TPicture;
  bmp, bmp1: TBitmap;
begin
  Pict:= TPicture.Create;
  Bmp:=Tbitmap.create;
  Bmp1:= Tbitmap.Create;
  try
    Case vartype(value) of
      varstring: begin
        Pict.LoadFromResourceName(HInstance, value);
        Bmp1.Assign(Pict.bitmap);
      end;
      else if not (ilIn=nil)then begin
        ilIn.GetBitmap(value, Bmp1);
      end;
    end;
    if twin then
    begin
      CropBitmap(Bmp1, Bmp, true);
      ilOut.AddMasked(Bmp, $FF00FF);
      // We need disabled image
      CropBitmap(Pict.Bitmap, Bmp, false);
      ilOut.AddMasked(Bmp, $FF00FF);
    end else ilOut.AddMasked(Bmp1, $FF00FF);
  except
     // Do nothing
  end;
  if Assigned(Bmp) then Bmp.Free;
  if Assigned(Bmp1) then Bmp1.Free;
  if Assigned(Pict) then Pict.Free;
end;


// Initialize buttons and menu images

procedure TFMailsInBox.InitButtons;
var
  i: integer;
begin
  // Create the array of buttons in order we need to place them in the bar
  // The order must meet the order for small buttons glyphs,
  // Change if we add or modify buttons;
  SetLength(BtnsArr, 13);
  BtnsArr[glImport].Btn:= BtnImport;
  BtnsArr[glAccountLog].Btn:= BtnAccountLog;
  BtnsArr[glGetAllMail].Btn:= BtnGetAllMail;
  BtnsArr[glGetAccMail].Btn:= BtnGetAccMail;
  BtnsArr[glLaunchClient].Btn:= BtnLaunchClient;
  BtnsArr[glDeleteAccount].Btn:= BtnDeleteAcc;
  BtnsArr[glAddAccount].Btn:= BtnAddAcc;
  BtnsArr[glEditAccount].Btn:= BtnEditAcc;
  BtnsArr[glSettings].Btn:= BtnSettings;
  BtnsArr[glLog].Btn:= BtnLog;
  BtnsArr[glHelp].Btn:= BtnHelp;
  BtnsArr[glAbout].Btn:= BtnAbout;
  BtnsArr[glQuit].Btn:= BtnQuit;
  // Initalise buttons record array
  for i:=0 to length(BtnsArr)-1 do
  begin
   BtnsArr[i].Bmp:= TBitmap.create;
   BtnsArr[i].Bmp.Assign(BtnsArr[i].Btn.Glyph);
   BtnsArr[i].Enabled:= false;
  end;
  // Large buttons. Copy buttons images in image list
  ILLargeBtns.clear;
  for i:=0 to length(BtnsArr)-1 do ILLargeBtns.AddMasked(BtnsArr[i].Btn.Glyph, $FF00FF) ;
  // Add mail clients defined images
  GetMnuImage(ILLargeBtns, 'LAUNCHOUTL', false);
  GetMnuImage(ILLargeBtns, 'LAUNCHTBIRD', false);
  GetMnuImage(ILLargeBtns, 'LAUNCHGMAIL', false);
  GetMnuImage(ILLargeBtns, 'LAUNCHWMAIL', false);
  GetMnuImage(ILLargeBtns, 'LAUNCHOUTLC', false);
  // Main Menu images
  GetMnuImage(ILMainMnu, glImport, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLaunchClient, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glQuit, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glGetAllMail, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glGetAccMail, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glSettings, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glAddAccount, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glEditAccount, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLog, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glHelp, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glAbout, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLaunchOutlook, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLaunchTbird, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLaunchGmail, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLaunchWin10mail, true, ILSmallBtns);
  GetMnuImage(ILMainMnu, glLaunchOutlkcom, true, ILSmallBtns);
  // Add tray menu images
  ILMnuTray.Clear;
  GetMnuImage(ILMnuTray, 'RESTORE16');
  GetMnuImage(ILMnuTray, 'MAXIMIZE16');
  GetMnuImage(ILMnuTray, 'ICONIZE16');
  GetMnuImage(ILMnuTray, glGetAllMail, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glAbout, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glQuit, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glLaunchClient, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glLaunchOutlook, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glLaunchTbird, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glLaunchGmail, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glLaunchWin10mail, true, ILSmallBtns);
  GetMnuImage(ILMnuTray, glLaunchOutlkcom, true, ILSmallBtns);
  // get Mails menu images
  ILMnuMails.Clear;
  GetMnuImage(ILMnuMails, 'MAILINFOS16');
  GetMnuImage(ILMnuMails, 'MAILANSWER16');
  GetMnuImage(ILMnuMails, 'MAILDELETE16');
  ILMnuAccounts.Clear;
  // get accounts menu images
  GetMnuImage(ILMnuAccounts, glAccountLog, true, ILSmallBtns);
  GetMnuImage(ILMnuAccounts, glGetAccMail, true, ILSmallBtns);
  GetMnuImage(ILMnuAccounts, glDeleteAccount, true, ILSmallBtns);
  GetMnuImage(ILMnuAccounts, glEditAccount, true, ILSmallBtns);
  GetMnuImage(ILMnuAccounts, 'ARROWUP16');
  GetMnuImage(ILMnuAccounts, 'ARROWDN16');

end;

//Dernière recherche il y a "days" jours ou plus ?

procedure TFMailsInBox.CheckUpdate(days: PtrInt);
var
  errmsg: string;
  sNewVer: string;
  CurVer, NewVer: int64;
  alertpos: TPosition;
  alertmsg: string;
begin
  //Dernière recherche il y a plus de 'days' jours ?
  errmsg := '';
  alertmsg:= '';
  if not visible then alertpos:= poDesktopCenter
  else alertpos:= poMainFormCenter;
  if (Trunc(Now)>Trunc(FSettings.Settings.LastUpdChk)+days) and (not FSettings.Settings.NoChkNewVer) then
  begin
     FSettings.Settings.LastUpdChk := Trunc(Now);
     AboutBox.Checked:= true;
     AboutBox.ErrorMessage:='';
     sNewVer:= AboutBox.ChkNewVersion;
     errmsg:= AboutBox.ErrorMessage;
     // Retry if nothing found the first time
     if (length(sNewVer)=0) and (length(errmsg)=0)then
     begin
       Application.ProcessMessages;
       //sNewVer:= AboutBox.ChkNewVersion;
       //errmsg:= AboutBox.ErrorMessage;
       sNewVer:= UpdateDlg.ChkNewVersion(UpdateDlg.ChkVerUrl, errmsg) ;
     end;
     if length(sNewVer)=0 then
     begin
       if length(errmsg)=0 then alertmsg:= sCannotGetNewVerList
       else alertmsg:= TranslateHttpErrorMsg(errmsg, HttpErrMsgNames);
       if AlertDlg(Caption,  alertmsg, [OKBtn, CancelBtn, sNoLongerChkUpdates],
                    true, mtError, alertpos)= mrYesToAll then FSettings.Settings.NoChkNewVer:= true;
       LogAddLine(-1, now, alertmsg);
       exit;
     end;
     NewVer := VersionToInt(sNewVer);
     // Cannot get new version
     if NewVer < 0 then exit;
     //CurVer := VersionToInt('0.1.0.0');     //Test version check
     CurVer := VersionToInt(version);
     if NewVer > CurVer then
     begin
       FSettings.Settings.LastVersion:= sNewVer;
       AboutBox.LUpdate.Caption := Format(AboutBox.sUpdateAvailable, [sNewVer]);
       LogAddLine(-1, now, AboutBox.LUpdate.Caption);
       AboutBox.NewVersion:= true;
       UpdateDlg.sNewVer:= sNewVer;
       UpdateDlg.NewVersion:= true;
       {$IFDEF WINDOWS}

       if UpdateDlg.ShowModal = mryes then MnuQuitClick(nil);    // New version install experimental
       {$ELSE}
         AboutBox.ShowModal;
       {$ENDIF}
     end else
     begin
       AboutBox.LUpdate.Caption:= AboutBox.sNoUpdateAvailable;
       LogAddLine(-1, now, AboutBox.sNoUpdateAvailable);
     end;
     FSettings.Settings.LastUpdChk:= now;
   end else
   begin
    if VersionToInt(FSettings.Settings.LastVersion)>VersionToInt(version) then
    begin
       AboutBox.LUpdate.Caption := Format(AboutBox.sUpdateAvailable, [FSettings.Settings.LastVersion]);
       AboutBox.NewVersion:= true;
    end else
    begin
      AboutBox.LUpdate.Caption:= AboutBox.sNoUpdateAvailable;
      // Already checked the same day
      if Trunc(FSettings.Settings.LastUpdChk) = Trunc(now) then AboutBox.checked:= true;
    end;
  end;
  // AboutBox.Translate(LangFile);
end;

// Initialize accounts base

procedure TFMailsInBox.LoadAccounts(filename: string);
var
  FilNamWoExt: string;
  i: integer;
begin
  FAccounts.Accounts.Reset;
  if not FileExists(AccountsFileName) then
  begin
    // Search if recent backup exists
    FilNamWoExt:= TrimFileExt(AccountsFileName);
    if FileExists(FilNamWoExt+'.bk0') then
    begin
      LogAddLine(-1, now, sRetConfBack);
      RenameFile(FilNamWoExt+'.bk0', AccountsFilename);
      for i := 1 to 5 do
        if FileExists(FilNamWoExt+'.bk' + IntToStr(i))
        // Renomme les précédentes si elles existent
        then
          RenameFile(FilNamWoExt+'.bk' + IntToStr(i),
            FilNamWoExt +'.bk' + IntToStr(i - 1));
    end else
    begin
      FAccounts.Accounts.SaveToXMLfile(filename);
      LogAddLine(-1, now, sCreNewConf);
    end;
  end;
  FAccounts.Accounts.LoadXMLfile(filename);
  if FAccounts.Accounts.Count=0 then
  begin
    BtnEditAcc.Enabled:= false;
  end;
  LogAddLine(-1, now, sLoadingAccounts);
end;

// Initializing stuff

procedure TFMailsInBox.Initialize;
var
  i: integer;
  defmailcli: string;
 // IniFile: TBbIniFile;
  tmplog: TstringList;
  CurAcc: TAccount;
  curcol: integer;
begin
  if initialized then exit;
  // Mail list images
  AccountPictures:= TPicture.Create;
  AccountPictures.LoadFromResourceName(HInstance, 'ACCOUNT216');
  MailPictures:= TPicture.Create;
  MailPictures.LoadFromResourceName(HInstance, 'MAILSTATES');
  // Tray icon
  TrayPicture:= Tpicture.Create;
  TrayPicture.LoadFromResourceName(HInstance, 'MAIL16');
  ChkMailImgCount:= ILChkMail.Count;
  // General pictures
  LaunchPicture:= Tpicture.Create;
  // Now, main settings
  FSettings.Settings.AppName:= LowerCase(ProgName);
  FAccounts.Accounts.AppName := LowerCase(ProgName);
  ConfigFileName:= MIBAppDataPath+'settings.xml';
  AccountsFileName:= MIBAppDataPath+'accounts.xml';
  FSettings.Settings.LangStr:= CurLangStr;
  LangFile:= TBbIniFile.Create(ExtractFilePath(Application.ExeName) + 'lang'+PathDelim+FSettings.Settings.LangStr+'.lng');
  Translate(LangFile);
  FSettings.Settings.ButtonBar:= true;
  LoadSettings(ConfigFileName);
  InitAboutBox;
  // In case of program's first use
  if length(FSettings.Settings.LastVersion)=0 then FSettings.Settings.LastVersion:= version;
  LoadAccounts(AccountsFileName);
  Application.Title:=Caption;
  if (Pos('64', OSVersion.Architecture)>0) and (OsTarget='32 bits') then
    MsgDlg(Caption, sUse64bit, mtInformation,  [mbOK], [OKBtn]);
  Application.ProcessMessages;
  // in case startup was done after a session end
  if FSettings.Settings.Restart then
  begin
    FSettings.Settings.Restart:= false;
    LogAddLine(-1, now, sRestart);
  end;
  {$IFDEF Linux}
     if not FSettings.Settings.Startup  then UnsetAutostart(ProgName);
  {$ENDIF}


   // Load last log file
  tmplog:= TStringList.Create;
  if FileExists(LogFileName) then
  begin
    tmpLog.LoadFromFile(LogFileName);
    MainLog:= tmpLog.text;
  end;
  BtnLaunchClient.Enabled:= not (FSettings.Settings.MailClient='');
  MnuLaunchClient.Enabled:= BtnLaunchClient.Enabled;
  // Accounts are sorted on their index
  FAccounts.Accounts.SortType:= cdcindex;
  FAccounts.Accounts.DoSort;
  PopulateAccountsList(false);
  if  FAccounts.Accounts.count > 0 then LVAccounts.ItemIndex:= 0
  else LVAccounts.ItemIndex:=  -1;
  FSettings.Settings.OnChange := @SettingsOnChange;
  FSettings.Settings.OnStateChange := @SettingsOnStateChange;
  //FSettings.Settings.OnQuitAlertChange:= @SettingsOnChange;
  FAccounts.Accounts.OnChange:= @AccountsOnChange;
  // Enumerate accounts, add names to log and update nextfire value
  for i:= 0 to FAccounts.Accounts.count-1 do
  begin
    // print accounts in log
    CurAcc:= FAccounts.Accounts.GetItem(i);
    if CurAcc.Enabled then
    AccountStatus:=Format('%s, Id : %u', [Format(sAccountEnabled,[CurAcc.Name]),CurAcc.UID])
    else AccountStatus:=Format('%s, Id : %u',[Format(sAccountDisabled, [CurAcc.Name]),CurAcc.UID]);
    LogAddLine(-1, now, AccountStatus);
  end;
  // Update infos as we may have changed nextfire time
  UpdateInfos;
  // OS version in Settings dialog status line
  FSettings.LStatus.Caption := OSVersion.VerDetail;
  // set small buttons only if asked
  if FSettings.Settings.SmallBtns then SetSmallBtns(bzSmall)
  else SetSmallBtns(bzNone); // customize only launchmail button
  Constraints.MinWidth:= BtnQuit.left+BtnQuit.width+10;
  if width < Constraints.MinWidth then width := Constraints.MinWidth;
  // Get default mail client and stores its name
  defmailcli:= FSettings.GetDefaultMailCllient;
  // Launch config dialog to set mail client and other stuff
  if length(FSettings.Settings.MailClient)=0 then
  begin
    FSettings.Settings.MailClient:= defmailcli;
    //BtnSettingsClick(self);
  end;
  LogAddLine(-1, now, 'Mail Client: '+FSettings.Settings.MailClientName);
  // TStringgrid MousetoCell give cell nearest the mouse click if
  // this property is true. To get -1 when mouse is outside a cell then
  SGMails.AllowOutboundEvents:=false;
  TrayMail.Hint:= sTrayHintNoMsg;
  if FSettings.Settings.StartupCheck then BtnGetAllMailClick(self);
  GetMailTimer.enabled:= true;
  //SettingsOnQuitAlertChange(self);
  // Set proper sort order and direction of mails display
  DisplayMails.SortType:= FSettings.Settings.MailSortTyp;
  DisplayMails.SortDirection:= FSettings.Settings.MailSortDir;
  CurCol:= 3;
  Case DisplayMails.SortType of
    cdcMessageFrom: CurCol:= 0;
    cdcMessageTo: CurCol:= 1;
    cdcMessageSubject: CurCol:= 2;
    cdcMessageDate: CurCol:= 3;
    cdcMessageSize: CurCol:= 4;
  end;
  SGMails.Columns[CurCol].Title.ImageIndex:= Ord(FSettings.Settings.MailSortDir)+1;
  Initialized:= true;
end;

procedure TFMailsInBox.InitAboutBox;
var
  IniFile: TBbIniFile;
begin
  // Check inifile with URLs, if not present, then use default
  IniFile:= TBbInifile.Create('mailsinbox.ini');
  AboutBox.ChkVerURL := IniFile.ReadString('urls', 'ChkVerURL','https://api.github.com/repos/bb84000/mailsinbox/releases/latest');
  AboutBox.UrlWebsite:= IniFile.ReadString('urls', 'UrlWebSite','https://www.sdtp.com');
  AboutBox.UrlProgSite:=  IniFile.ReadString('urls', 'UrlProgSite','https://github.com/bb84000/mailsinbox/wiki');
  AboutBox.UrlSourceCode:= IniFile.ReadString('urls', 'UrlSourceCode','https://github.com/bb84000/mailsinbox');
  UpdateDlg.ChkVerUrl:=  AboutBox.ChkVerURL;
  UpdateDlg.UrlInstall:= IniFile.ReadString('urls', 'UrlInstall', 'https://github.com/bb84000/mailsinbox/raw/refs/heads/master/mailsinbox.zip');
  UpdateDlg.ExeInstall:= IniFile.ReadString('urls', 'ExeInstall', 'InstallMailsInBox.exe');
  ChkVerInterval:= IniFile.ReadInt64('urls', 'ChkVerInterval', 3);
  if Assigned(IniFile) then IniFile.free;
  // Language dependent variables are updated in ModLangue procedure
  AboutBox.Width:= 400; // to have more place for the long product name
  AboutBox.Image1.Picture.LoadFromResourceName(HInstance, 'ABOUTIMG');
  AboutBox.LCopyright.Caption := GetVersionInfo.CompanyName + ' - ' + DateTimeToStr(CompileDateTime);
  AboutBox.LVersion.Caption := 'Version: ' + Version + ' (' + OS + OSTarget + ')';
  AboutBox.LUpdate.Hint := AboutBox.sLastUpdateSearch + ': ' + DateToStr(FSettings.Settings.LastUpdChk);
  AboutBox.Version:= Version;
  AboutBox.ProgName:= ProgName;
  AboutBox.LastUpdate:= FSettings.Settings.LastUpdChk;
  AboutBox.LastVersion:= FSettings.Settings.LastVersion;
  AboutBox.autoUpdate:= true;
  // Populate UpdateBox with proper variables
  UpdateDlg.ProgName:= ProgName;
  UpdateDlg.NewVersion:= false;
  AboutBox.Translate(LangFile);
  UpdateDlg.Translate(LangFile);
end;

// Change BtnQuit behaviour: one click if alert enabled,
// double click if aleret disabled
// procedure to call on each NoQuitAlert setting change
// No longer used, see BtnQuitMouseDown procedure

{procedure TFMailsInBox.SettingsOnQuitAlertChange(sender: TObject);
begin
  if FSettings.Settings.NoQuitAlert then
  begin
    BTnQuit.OnDblClick:= @BtnQuitDblClick;
    BTnQuit.OnClick:= nil;
  end else
  begin
    BTnQuit.OnDblClick:= nil;
    BTnQuit.OnClick:= @BtnQuitClick;
  end;
end; }

// Hide taskbar icon, App is only in tray

function TFMailsInBox.HideOnTaskbar: boolean;
begin
  result:= false;
  if (WindowState=wsMinimized) and FSettings.Settings.HideInTaskbar then
  begin
    result:= true;
    visible:= false;
  end;
end;

// Change of form size or splitters position

procedure TFMailsInBox.DoChangeBounds(Sender: TObject);
begin
  // Change time display position and status label width when window width change
  if Sender= self then begin
    LNow.left:= Clientwidth-LNow.width-10;
    LStatus.width:= Clientwidth -LNow.width-20;
  end;
  SettingsChanged:= FSettings.Settings.SavSizePos;
end;

// change size of buttons in toolbar

procedure TFMailsInBox. SetSmallBtns(small:TBtnSize);
var
  i: integer;
  imgndx: integer;
  btnsiz: integer;
begin
  imgndx:= glLaunchClient;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'outlook') then imgndx:= glLaunchOutlook;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'thunder') then imgndx:= glLaunchTbird;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'gmail') then imgndx:= glLaunchGmail;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'windows') then imgndx:= glLaunchWin10mail;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'outlook.com') then imgndx:= glLaunchOutlkcom;
  Case Small of
    bzSmall: begin
      PnlToolbar.height:= 34;
      btnsiz:= 24;
    end;
    bzLarge: begin
      PnlToolbar.height:= 42;
      btnsiz:= 32;
    end;
  end;
  if not (small=bzNone) then
    for i:=0 to length(BtnsArr)-1 do
    begin
      BtnsArr[i].Btn.Height:= btnsiz;
      BtnsArr[i].Btn.Width:= btnsiz;
      BtnsArr[i].Btn.Left:= 10+(i)*(btnsiz+8);
      if Small=bzSmall then ILSmallBtns.GetBitmap(i, BtnsArr[i].Btn.Glyph)
      else ILLargeBtns.GetBitmap(i, BtnsArr[i].Btn.Glyph);
    end;
  // customize mail client image
  if not FSEttings.Settings.SmallBtns then              // large buttons
  begin
    if imgndx>glQuit then ILLargeBtns.GetBitmap(imgndx, BtnLaunchClient.Glyph)
    else ILLargeBtns.GetBitmap(Ord(glLaunchClient), BtnLaunchClient.Glyph);
  end else
  begin
    if imgndx>glQuit then ILSmallBtns.GetBitmap(imgndx, BtnLaunchClient.Glyph)
    else ILSmallBtns.GetBitmap(Ord(glLaunchClient), BtnLaunchClient.Glyph);
  end ;
  if length(FSettings.Settings.MailClientName)=0 then BtnLaunchClient.Hint:= sBtnLaunchClientDef
  else BtnLaunchClient.Hint:= Format(sBtnLaunchClientCust, [FSettings.Settings.MailClientName]);
end;

// Load configuration and database from file

procedure TFMailsInBox.LoadSettings(Filename: string);
var
  winstate: TWindowState;
  i: integer;
begin
  with FSettings do
  begin
    LogAddLine(-1, now, sLoadConf);
    Settings.LoadXMLFile(Filename);
    self.Position:= poDesktopCenter;
    if Settings.SavSizePos then
    try
      WinState := TWindowState(StrToInt('$' + Copy(Settings.WState, 1, 4)));
      self.Top := StrToInt('$' + Copy(Settings.WState, 5, 4));
      self.Left := StrToInt('$' + Copy(Settings.WState, 9, 4));
      self.Height := StrToInt('$' + Copy(Settings.WState, 13, 4));
      self.Width := StrToInt('$' + Copy(Settings.WState, 17, 4));
      self.PnlLeft.width:= StrToInt('$' + Copy(Settings.WState, 21, 4));
      self.PnlAccounts.Height:= StrToInt('$' + Copy(Settings.WState, 25, 4));
      For i:= 0 to 4 do self.SGMails.Columns[i].Width:= StrToInt('$'+Copy(Settings.WState,29+(i*4),4)) ;
      FLogView.Top:= StrToInt('$' + Copy(Settings.WState, 49, 4));
      FLogView.Left:= StrToInt('$' + Copy(Settings.WState, 53, 4));
      FLogView.Height:= StrToInt('$' + Copy(Settings.WState, 57, 4));
      FLogView.Width:= StrToInt('$' + Copy(Settings.WState, 61, 4));;
      ;self.WindowState := WinState;
      PrevLeft:= self.left;
      PrevTop:= self.top;
      if Winstate = wsMinimized then
      begin
        //Application.Minimize;
        StartMini:= true;
      end else self.WindowState := WinState;
    except
    end;
    // Display buttons bar and/or menu bar
    if Settings.MenuBar then
    begin
      self.Menu:= MainMnu;
    end else
    begin
      self.Menu:= nil;
    end;
    PnlToolbar.visible:= Settings.ButtonBar;
    if not PnlToolbar.visible then self.Menu:= MainMnu;
    MnuDisplayMenu.checked:= (self.Menu=MainMnu);
    MMnuDisplayMenu.Checked:= MnuDisplayMenu.checked;
    MnuDisplayBar.Checked:= PnlToolbar.visible;
    MMnuDisplayBar.checked:= MnuDisplayBar.Checked;
    // Get columns width to use at application close
    sColumnswidth:='';
    For i:= 0 to 4 do  sColumnswidth:= sColumnswidth+IntToHex(self.SGMails.Columns [i].Width, 4);
    if settings.StartMini then
    begin
      // On lazarus 2.2.2 and over minimize at startup no longer works properly
      // Application.minimize let a minmized window on the desktop when called in activate event
      // Send an user message at the end of activation procedure to minimize once the form is shown
      StartMini:= true;
     end;
    // Détermination de la langue (si pas dans settings, langue par défaut)
    if Settings.LangStr = '' then Settings.LangStr := CurLangStr;
    try
      FindAllFiles(LangNums, ExtractFilePath(Application.ExeName) + 'lang', '*.lng', true); //find all language files
      if LangNums.count > 0 then
      begin
        for i:= 0 to LangNums.count-1 do
        begin
          LangFile:= TBbInifile.Create(LangNums.Strings[i]);
          LangNums.Strings[i]:= TrimFileExt(ExtractFileName(LangNums.Strings[i]));
          FSettings.CBLangue.Items.Add(LangFile.ReadString('common', 'Language', 'Inconnu'));
          if LangNums.Strings[i] = Settings.LangStr then LangFound := True;
        end;
      end;
    except
      LangFound := false;
    end;
    // Si la langue n'est pas traduite, alors on passe en Anglais
    if not LangFound then
    begin
      Settings.LangStr := 'en';
    end;
  end;
  LangFile:= TBbIniFile.Create(ExtractFilePath(Application.ExeName) + 'lang'+PathDelim+FSettings.Settings.LangStr+'.lng');
  Translate(LangFile);
  SettingsChanged := false;
end;

// Procedure used during QueryEndSession and Formclose function
procedure TFMailsInBox.BeforeClose;
var
  s: string;
  i: integer;
  curcolsw: string;
begin
  // check if columns width has changed
  curcolsw:='';
  For i:= 0 to 4 do  curcolsw:= curcolsw+IntToHex(self.SGMails.Columns [i].Width, 4);
  if (curcolsw <> sColumnswidth) then DoChangeBounds(self);
  if FSettings.Settings.Startup then SetAutostart(progname, Application.exename)
      else UnSetAutostart(progname);
  if AccountsChanged then
  begin
    SaveConfig(All);
  end else
  begin
    if SettingsChanged then SaveConfig(Setting) ;
  end;
  LogAddLine(-1, now, sClosingProg);
  LogAddLine(-1, now, '************************');
  if FSettings.Settings.SaveLogs then
  begin
    s:= Mainlog+SessionLog.text;
    SessionLog.text:=s;
  end;
  SessionLog.SaveToFile(LogFileName);
  Application.ProcessMessages;
end;

procedure TFMailsInBox.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if CanClose then
  begin
     CloseAction:= caFree;
    if CheckingMail then
    begin
      ShowMessage(sCannotQuit);
      CloseAction:= caNone;
    end else
    begin
      // check if something has changed in accounts and settings
      BeforeClose;
      CloseAction:= caFree;
    end;
  end else
  begin
    // Alertbox only if not disabled in settings
    if not FSettings.Settings.NoCloseAlert then
    begin
       Case AlertDlg(Caption, sNoCloseAlert, [OKBtn, CancelBtn, sAlertBoxCBNoShowAlert], true) of
         mrOk: MnuIconizeClick(sender);
         mrYesToAll: begin
            FSettings.Settings.NoCloseAlert:= true;
            SettingsChanged:= true;
            LogAddLine(-1, now, sNoShowCloseAlert);
            MnuIconizeClick(sender);
         end;
       end;
    end else MnuIconizeClick(sender);
    CloseAction := caNone;
  end;
end;


// Save configuration and database to file

function TFMailsInBox.SaveConfig(typ: TSaveMode): boolean;
var
  FilNamWoExt: string;
  i: integer;
begin
  Result := False;
  if (Typ= Setting) or (Typ = All) then
  with FSettings do
  begin
    Settings.WState:= '';
    if self.Top < 0 then self.Top:= 0;
    if self.Left < 0 then self.Left:= 0;
    // Main form size and position
    Settings.WState:= IntToHex(ord(self.WindowState), 4)+IntToHex(self.Top, 4)+
                      IntToHex(self.Left, 4)+IntToHex(self.Height, 4)+IntToHex(self.width, 4)+
                      IntToHex(self.PnlLeft.width, 4)+
                      IntToHex(self.PnlAccounts.height, 4);
    // Mail list columns size
    For i:= 0 to 4 do Settings.WState:= Settings.WState+IntToHex(self.SGMails.Columns [i].Width, 4);

    // Log window size
    Settings.WState:= Settings.WState+IntToHex(FLogView.Top, 4)+IntToHex(FLogView.Left, 4)+
                                      IntToHex(FLogView.Height, 4)+IntToHex(FLogView.width, 4);
    Settings.Version:= version;
    if FileExists (AccountsFileName) then
    begin
      if (Typ = All) then
      begin
        // On sauvegarde les versions précédentes parce que la base de données a changé
        FilNamWoExt:= TrimFileExt(AccountsFileName);
        if FileExists (FilNamWoExt+'.bk5')                   // Efface la plus ancienne
        then  DeleteFile(FilNamWoExt+'.bk5');                // si elle existe
        For i:= 4 downto 0
        do if FileExists (FilNamWoExt+'.bk'+IntToStr(i))     // Renomme les précédentes si elles existent
           then  RenameFile(FilNamWoExt+'.bk'+IntToStr(i), FilNamWoExt+'.bk'+IntToStr(i+1));
        RenameFile(AccountsFileName, FilNamWoExt+'.bk0');
        FAccounts.Accounts.SaveToXMLfile(AccountsFileName);
      end;
      // la base n'a pas changé, on ne fait pas de backup
      FSettings.settings.SaveToXMLfile(ConfigFileName);
    end else
    begin
      FAccounts.Accounts.SaveToXMLfile(AccountsFileName);
      settings.SaveToXMLfile(ConfigFileName); ;
    end;
    result:= true;
  end;
end;

// Create or update accounts list and associated stuff
procedure TFMailsInBox.PopulateAccountsList (notify: boolean);
var
  Listitem: TlistItem;
  i,j: Integer;
  AccBmp : TBitmap;
  CurAcc: TAccount;
  NewMsgsCnt: integer;
  sTmpHint: string;
  sTrayNewHint: string;
  sTrayBallHint: string;
  sLineEnd: string;
  totalNewMsgs: integer;
Begin
  if FAccounts.Accounts.Count = 0 then
  begin
    BtnGetAllMail.Enabled:= false;
    MMnuGetAllMails.Enabled:= BtnGetAllMail.Enabled;
    BtnGetAccMail.Enabled:= false;
    MMnuGetAccMails.Enabled:= BtnGetAccMail.Enabled;
    BtnDeleteAcc.Enabled:= false;
    BtnEditAcc.Enabled:= false;
    MMnuEditAcc.Enabled:= BtnEditAcc.Enabled;
    LVAccounts.Clear;
    LVAccounts.PopupMenu:= nil;
    exit;
  end;
  LVAccounts.PopupMenu:= MnuAccount;;
  BtnGetAllMail.Enabled:= true;
  MMnuGetAllMails.Enabled:= BtnGetAllMail.Enabled;
  BtnGetAccMail.Enabled:= true;
  MMnuGetAccMails.Enabled:= BtnGetAccMail.Enabled;
  BtnDeleteAcc.Enabled:= true;
  BtnEditAcc.Enabled:= true;
  MMnuEditAcc.Enabled:= BtnEditAcc.Enabled;
  sTrayNewHint:='';
  sTrayBallHint:='';
  TrayMail.Hint:= '';
  TrayMail.BalloonHint:='';
  sLineEnd:='';
  totalNewMsgs:=0;
  AccBmp:= TBitmap.create;
  Application.ProcessMessages;  //23/3/25
  LVAccounts.Clear;
  if Assigned(LVAccounts.SmallImages) then LVAccounts.SmallImages.Clear;
  ILTray.Clear;
  TrayPicture.LoadFromResourceName(HInstance, 'MAIL16');
  Application.ProcessMessages;  //23/3/25
  ILTray.AddMasked(TrayPicture.Bitmap, $FF00FF); // Reset after each mail checking
  for i := 0 to FAccounts.Accounts.Count-1 do
  Try
    NewMsgsCnt:=0;
    stmpHint:='';
    CurAcc:= FAccounts.Accounts.GetItem(i);
    if CurAcc.Enabled then CropBitmap(AccountPictures.Bitmap, AccBmp, true)
    else CropBitmap(AccountPictures.Bitmap, AccBmp, false)  ;
    ListItem := LVAccounts.items.add;  // prépare l'ajout
    if CurAcc.Mails.count > 0 then
    begin
      DrawTheIcon(AccBmp, CurAcc.Mails.count, CurAcc.Color  );
      Application.ProcessMessages;  //23/3/25
      DrawTheIcon(TrayPicture.Bitmap, CurAcc.Mails.count, CurAcc.Color  );
      Application.ProcessMessages;  //23/3/25
      ILTray.AddMasked(TrayPicture.Bitmap, $FF00FF);  // modified icon
    end;
    if CurAcc.Error then
    begin
      DrawTheIcon(AccBmp, -1, CurAcc.Color);
      Application.ProcessMessages;  //23/3/25
      DrawTheIcon(TrayPicture.Bitmap, -1, CurAcc.Color);
      Application.ProcessMessages;  //23/3/25
      ILTray.AddMasked(TrayPicture.Bitmap, $FF00FF);  // modified icon
    end;
    LVAccounts.SmallImages.AddMasked(AccBmp,$FF00FF);
    ListItem.ImageIndex := i;
    Listitem.Caption :=  CurAcc.Name;    // ajoute le nom
    if CurAcc.Mails.count >0 then
    begin
      for j:= 0 to CurAcc.Mails.count-1 do
        if CurAcc.Mails.GetItem(j).MessageNew then
        begin
          Inc(NewMsgsCnt);
          // if new message already displayed but still new message
          // mark it to not fire ballon for it
          if not CurAcc.Mails.GetItem(j).MessageDisplayed then
          begin
            CurAcc.Mails.ModifyField(j, 'MessageDisplayed', true);
            Inc(totalNewMsgs);
          end;
        end;
       if CurAcc.Mails.count=1 then sTmpHint:=sTrayHintMsg else sTmpHint:=sTrayHintMsgs;
      sTmpHint:=Format(sTmpHint, [CurAcc.Name, CurAcc.Mails.count]);
      if NewMsgsCnt>0 then
      begin
        if NewMsgsCnt=1 then
        begin
          sTrayBallHint:= sTrayBallHintMsg;
          sTrayNewHint:= sTrayHintNewMsg;
        end else
        begin
          sTrayBallHint:=sTrayBallHintMsgs;
          sTrayNewHint:=sTrayHintNewMsgs;
        end;
        if i<(FAccounts.Accounts.Count-1) then sLineEnd:=#10 else sLineEnd:='';
        TrayMail.BalloonHint:= TrayMail.BalloonHint+Format(sTrayBallHint, [CurAcc.Name,NewMsgsCnt,sLIneEnd]);
        sTmpHint:=Format(sTmpHint, [Format(sTrayNewHint, [NewMsgsCnt])])
      end else sTmpHint:= Format(sTmpHint, ['']);
     TrayMail.Hint:=TrayMail.Hint+sTmpHint+#10;
    end;
  Except
    // Error in process
  end;
  if TrayMail.Hint='' then TrayMail.Hint:= sTrayHintNoMsg;
  LVAccounts.ItemIndex:= 0;
  if Assigned(AccBmp) then AccBmp.free;
  if FSettings.Settings.RestNewMsg and not visible then MnuRestoreClick(self);
  // Ony show notification in  case of real new mail
  if FSettings.Settings.Notifications and notify and (length(TrayMail.BalloonHint)>0)
     and (totalNewMsgs>0)  then TrayMail.ShowBalloonHint;
end;


// Event fired by any change of settings values

procedure TFMailsInBox.SettingsOnChange(Sender: TObject);
begin
SettingsChanged := True;
end;

// Event fired by any state change (window state and position)

procedure TFMailsInBox.SettingsOnStateChange(Sender: TObject);
begin
  SettingsChanged := True;
end;

// Event fired by CurAcc change

procedure TFMailsInBox.AccountsOnChange(Sender: TObject);
begin
  AccountsChanged := True;
end;

// edit or add an CurAcc

procedure TFMailsInBox.BtnEditAccClick(Sender: TObject);
var
  CurAcc: TAccount;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  with FAccounts do
  begin
    if (TSpeedButton(Sender).Name='BtnAddAcc') or (TMenuItem(Sender).Name='MMnuAddAcc') then
    begin
      CurAcc:= Default(TAccount);
      CurAcc.Name:= sNewAccount;
      CurAcc.UID:=0;
      Caption:= BtnAddAcc.Hint;
    end;
    if ((TSpeedButton(Sender).Name='BtnEditAcc') or (TMenuItem(Sender).Name='MnuEditAcc')
         or (TMenuItem(Sender).Name='MMnuEditAcc')) then
    begin
       if ndx <0 then
       begin
         ShowMessage(Format(sPlsSelectAcc, [sToEditAcc]));
       exit;
    end;
      CurAcc:= Accounts.GetItem(ndx);
      Caption:= BtnEditAcc.Hint;
    end;
    EName.Text:= CurAcc.Name;
    EServer.Text:= CurAcc.Server;
    CBProtocol.ItemIndex:=ord(CurAcc.Protocol);
    EUserName.Text:= CurAcc.UserName;
    EPassword.Text:= CurAcc.Password;
    EEmail.Text:= CurAcc.Email;
    CBSSL.ItemIndex:= CurAcc.SSL;
    EPort.Text:= IntToStr(CurAcc.Port);
    CBSecureAuth.Checked:= CurAcc.SecureAuth;
    EInterval.Value:= CurAcc.Interval;
    EReplyEmail.Text:= CurAcc.ReplyEmail;
    CBEnabledAcc.Checked:= CurAcc.Enabled;
    CBColorAcc.Selected:= CurAcc.Color;
    ESoundFile.Text:= CurAcc.SoundFile;
    BtnPlaySound.Enabled:=not (length(ESoundFile.Text)=0);
    if ShowModal=mrOK then
    begin
      CurAcc.Name:= EName.Text;
      CurAcc.Server:= EServer.Text;
      CurAcc.protocol:= TProtocols(CBProtocol.ItemIndex);
      CurAcc.UserName:= EUserName.Text;
      CurAcc.Password:= EPassword.Text;
      CurAcc.Email:= EEmail.Text;
      CurAcc.SSL:= CBSSL.ItemIndex;
      CurAcc.Port:= StringToInt(EPort.Text);
      CurAcc.SecureAuth:= CBSecureAuth.Checked;
      CurAcc.Interval:= EInterval.Value;
      CurAcc.ReplyEmail:= EReplyEmail.Text;
      CurAcc.Enabled:= CBEnabledAcc.Checked;
      CurAcc.Color:= CBColorAcc.Selected;
      if (TSpeedButton(Sender).Name='BtnEditAcc') and  (ndx>=0) then
      Accounts.ModifyAccount(ndx, CurAcc)
      else Accounts.AddAccount(CurAcc);
      PopulateAccountsList (false);
      if (TSpeedButton(Sender).Name='BtnEditAcc') then
      begin
        LVAccounts.ItemIndex:= ndx;
        LogAddLine(CurAcc.UID, now, Format(sAccountChanged, [CurAcc.Name]));
      end else
      begin
        LVAccounts.ItemIndex:= LVAccounts.Items.count-1 ;
        LogAddLine(Accounts.GetItem(LVAccounts.ItemIndex).UID , now, Format(sAccountAdded, [CurAcc.Name]));
      end;
      if AccountsChanged then SaveConfig(All);
      AccountsChanged:= False;
    end;
  end;
end;

procedure TFMailsInBox.BtnCloseClick(Sender: TObject);
begin
   close;
end;

procedure TFMailsInBox.FormDestroy(Sender: TObject);
var
  i: integer;
begin
  if Assigned(TrayTimerBmp) then TrayTimerBmp.free;
  if Assigned(LangNums) then LangNums.free;
  if Assigned(LangFile) then LangFile.free;
  for i:=0 to length(BtnsArr)-1 do if Assigned(BtnsArr[i].Bmp) then BtnsArr[i].Bmp.free;
  //if Assigned(ChkMailTimer) then ChkMailTimer.Destroy;
  if Assigned(SessionLog) then SessionLog.free;
  if Assigned(AccountPictures) then AccountPictures.Free;
  if Assigned(TrayPicture) then TrayPicture.Free;
  if Assigned(MailPictures) then MailPictures.Free;
  if Assigned(LaunchPicture) then LaunchPicture.Free;
  if Assigned(slLastFires) then slLastFires.Free;
  if Assigned(slNextFires) then slNextFires.Free;
  if Assigned(DisplayMails) then DisplayMails.free;
end;

procedure TFMailsInBox.FormWindowStateChange(Sender: TObject);
begin

end;

// Timer firing periodic mail checking

procedure TFMailsInBox.GetMailTimerTimer(Sender: TObject);
var
  i: integer;
  min: TDateTime;
  CurAcc: TAccount;
   ndx: integer;
begin
  if FAccounts.Accounts.Count=0 then exit;
  ndx:= LVAccounts.ItemIndex;
  if ndx <0 then exit;
  for i:=0 to FAccounts.Accounts.count-1 do
  begin
    // current CurAcc is enabled and interval defined
    CurAcc:= FAccounts.Accounts.GetItem(i);
    if CurAcc.Enabled and (CurAcc.Interval>0) and (now> GetFire(CurAcc, fmNext)) then
    begin
      min:= EncodeTime(0,CurAcc.interval,0,0);
      GetAccMail(i);
      SetFire(CurAcc, now+min, fmNext);
    end;
  end;
  LVAccounts.ItemIndex:=ndx;
end;

// POP3 connected event

procedure TFMailsInBox.Id_client_Connected(Sender: TObject);
var
  Curacc: TAccount;
begin
  CurAcc:= FAccounts.Accounts.GetItem(CurAccPend);
  LStatus.Caption:= Format(sConnectedToServer, [Curacc.Name, CurAcc.Server]);
  LogAddLine(CurAcc.UID, now, LStatus.Caption );
end;

procedure TFMailsInBox.Id_client_Disconnected(Sender: TObject);
var
  Curacc: TAccount;
begin
  CurAcc:= FAccounts.Accounts.GetItem(CurAccPend);
  LStatus.Caption:= Format(sDisconnectedServer, [Curacc.Name, CurAcc.Server]);
  LogAddLine(CurAcc.UID, now, LStatus.Caption );
end;

procedure TFMailsInBox.IdPOP3_1Status(ASender: TObject;
  const AStatus: TIdStatus; const AStatusText: string);
begin
end;

// IAccount selection in account list
// use Item and selected parameters

procedure TFMailsInBox.LVAccountsSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  AccName: string;
 // ndx: integer;
begin
  //ndx:= LVAccounts.ItemIndex;
  //if ndx >= 0 then
  if (Selected and (Item.index>=0))then
  begin
    //BtnGetAccMail.Enabled:= FAccounts.Accounts.GetItem(ndx).enabled;
    BtnGetAccMail.Enabled:= FAccounts.Accounts.GetItem(Item.Index).enabled;
    MnuGetAccMail.Enabled:= BtnGetAccMail.Enabled;
    LVAccounts.PopupMenu:= MnuAccount;
    //AccName:= FAccounts.Accounts.GetItem(ndx).Name;      // assigned to Item.Caption
    AccName:= Item.Caption;
    BtnAccountLog.Hint:= Format(sBtnLogHint, [AccName]);
    MnuAccountLog.Caption:= BtnAccountLog.Hint;
    BtnGetAccMail.Hint:= Format(sBtnGetAccMailHint, [AccName]);
    MnuGetAccMail.Caption:= BtnGetAccMail.Hint;
    BtnDeleteAcc.Hint:= Format(sBtnDeleteHint, [AccName]);
    MnuDeleteAcc.Caption:= BtnDeleteAcc.Hint;
    BtnEditAcc.Hint:= Format(sBtnEditAccHint, [AccName]);
    MnuEditAcc.Caption:= BtnEditAcc.Hint;
    UpdateInfos;
  end else LVAccounts.PopupMenu:= nil;
  //PopulateMailsList(ndx);
  PopulateMailsList(Item.Index);
end;

procedure TFMailsInBox.MMnuClick(Sender: TObject);
var
  imgndx: integer;
begin
  imgndx:= 2;
  MMnuImport.Caption:= BtnImport.Hint;
  if MMnuImport.Enabled then MMnuImport.ImageIndex:= 0 else MMnuImport.ImageIndex:=1;
  MMnuLaunchClient.Caption:= MnuLaunchClient.Caption;
  MMmnuQuit.Caption:= MnuQuit.Caption;
  if MMmnuQuit.Enabled then MMmnuQuit.ImageIndex:= 4 else MMmnuQuit.ImageIndex:= 5;
  MMnuGetAllMails.Caption:= BtnGetAllMail.Hint;
  if MMnuGetAllMails.Enabled then MMnuGetAllMails.ImageIndex:= 6 else MMnuGetAllMails.ImageIndex:= 7;
  MMnuGetAccMails.Caption:= BtnGetAccMail.Hint;
  if MMnuGetAccMails.Enabled then MMnuGetAccMails.ImageIndex:= 8 else MMnuGetAccMails.ImageIndex:= 9;
  MMnuGetAccMails.Enabled:= BtnGetAccMail.Enabled;
  MMnuSettings.Caption:= BtnSettings.Hint;
  if MMnuSettings.Enabled then MMnuSettings.ImageIndex:= 10 else MMnuSettings.ImageIndex:= 11;
  MMnuAddAcc.Caption:= BtnAddAcc.Hint;
  if MMnuAddAcc.Enabled then MMnuAddAcc.ImageIndex:= 12 else MMnuAddAcc.ImageIndex:= 13;
  MMnuEditAcc.Caption:= BtnEditAcc.Hint;
  if MMnuEditAcc.Enabled then MMnuEditAcc.ImageIndex:= 14 else MMnuEditAcc.ImageIndex:= 15;
   MMnuLog.Caption:= BtnLog.Hint;
  if MMnuLog.Enabled then MMnuLog.ImageIndex:= 16 else MMnuLog.ImageIndex:= 17;
  MMnuHelp.Caption:= BtnHelp.Hint;
  if MMnuHelp.Enabled then MMnuHelp.ImageIndex:= 18 else MMnuHelp.ImageIndex:= 19;
  MMnuAbout.Caption:= BtnAbout.Hint;
  if MMnuAbout.Enabled then MMnuAbout.ImageIndex:= 20 else MMnuAbout.ImageIndex:= 21;
    // Change images for mail client
  if AnsiContainsText(FSettings.Settings.MailClientName, 'outlook') then imgndx:= 22;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'thunder') then imgndx:= 24;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'gmail') then imgndx:= 26;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'windows') then imgndx:= 28;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'outlook.com') then imgndx:= 30;
  if MMnuLaunchClient.Enabled then MMnuLaunchClient.ImageIndex:=imgndx else MMnuLaunchClient.ImageIndex:=imgndx+1;
end;


function TFMailsInBox.GetFire(CurAcc: Taccount; mode: TFireMode): TDateTime;
var
  uidfnd: boolean;
  A: TStringArray;
  i: integer;
  cnt: integer;
  s: string;
begin
  uidfnd:= false;
  result:= 0;
  if mode=fmLast then cnt:= FSettings.Settings.LastFires.count
  else cnt:= FSettings.Settings.NextFires.count;
  if cnt> 0 then
  begin
    for i:= 0 to cnt-1 do
    begin
      if mode=fmlast then s:= FSettings.Settings.LastFires.Strings[i]
      else s:= FSettings.Settings.NextFires.Strings[i];
      if pos(InttoStr(CurAcc.UID), s) >0 then
      begin
        A:= s.split('|');
        uidfnd:= true;
        break;
      end;
    end;
    if uidfnd then result:= UnixToDateTime(StrToInt64Def(A[1], 0));
  end;

end;

procedure TFMailsInBox.UpdateInfos;
var
  ndx: integer;
  msgs: integer;
  msgsfnd: string;
  CurAcc: TAccount;
  slastfire: string;
  dt:TDateTime;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx >= 0 then
  begin
    CurAcc:= FAccounts.Accounts.GetItem(ndx);
    msgs:= CurAcc.Mails.Count;
    dt:= GetFire(CurAcc, fmLast);
    if dt= 0 then slastfire:= sNeverChecked
    else slastfire:= TimeDateToString(dt);
    RMInfos.Clear;
    if CurAcc.Enabled then RMInfos.Lines.Add(Format(sAccountEnabled, [CurAcc.Name]))
    else RMInfos.Lines.Add(Format(sAccountDisabled, [CurAcc.Name]));
    RMInfos.Lines.Add(Format(sEmailCaption, [CurAcc.Email]));
    RMInfos.Lines.Add(Format(sLastCheckCaption, [slastfire]));
    if CurAcc.Enabled then
    begin
      if CurAcc.error then RMInfos.Lines.Add(CurAcc.ErrorStr);

      if msgs>1 then msgsfnd:= Format(sMsgsFound, [msgs])
      else msgsfnd:= Format(sMsgFound, [msgs]);
      RMInfos.Lines.Add(msgsfnd);
      LStatus.Caption:= Format(sLStatusCaption, [msgsfnd, CurAcc.Name, slastfire]);
      RMInfos.Lines.add(Format(sNextCheckCaption, [TimeDateToString(GetFire(CurAcc, fmNext))]));
    end else
    begin
      LStatus.Caption:= Format(sAccountDisabled, [CurAcc.Name]);
    end;
   end;
end;

procedure TFMailsInBox.PopulateMailsList(index: integer);
var
  i, j: integer;
  siz: integer;
  CurAcc: TAccount;
  s: string;
  oldUIDL: string;
begin
  // remember uidl of previous selected message  , stored in aMailsList
  if (SGMails.Row>0) and (SGMails.Row<=length(aMailsList)) then  oldUIDL:= aMailsList[SGMails.row-1]
  else oldUIDL:= '';
  SGMails.RowCount:=1;
  DisplayMails.Reset;
  // If we display only selected accout messages
  if not FSettings.Settings.DisplayAllAccMsgs then
  begin
    if index<0 then exit;
    CurAcc:= FAccounts.Accounts.GetItem(index);
    if (CurAcc.Mails.count>0) then
      for j:=0 to CurAcc.Mails.count-1 do DisplayMails.AddMail(CurAcc.Mails.GetItem(j));
  end else
  begin
    For i:=0 to FAccounts.Accounts.count-1 do
    begin
      CurAcc:= FAccounts.Accounts.GetItem(i);
      if (CurAcc.Mails.count> 0) then
      begin
        for j:= 0 to CurAcc.Mails.count-1 do
          DisplayMails.AddMail(CurAcc.Mails.GetItem(j));
      end;
    end;
  end;
  if DisplayMails.count=0 then exit;
  DisplayMails.SortType:= FSettings.Settings.MailSortTyp;
  DisplayMails.SortDirection:= FSettings.Settings.MailSortDir;
  // array of mails uidl to retreive later previous selected mail
  SetLength(aMailsList, DisplayMails.count);
  j:= DisplayMails.count+1;
  SGMails.RowCount:= j;
  for i:=0 to DisplayMails.Count-1 do
  begin
    SGMails.Cells[0,i+1]:= DisplayMails.GetItem(i).MessageFrom;
    SGMails.Cells[1,i+1]:= DisplayMails.GetItem(i).AccountName;
    SGMails.Cells[2,i+1]:= DisplayMails.GetItem(i).MessageSubject;
    SGMails.Cells[3,i+1]:= TimeDateToString(DisplayMails.GetItem(i).MessageDate);
    aMailsList[i]:= DisplayMails.GetItem(i).MessageUIDL;
    // Change unit with size value
    siz:= DisplayMails.GetItem(i).MessageSize;
    if siz<20480 then s:= InttoStr(siz)+' '+sBytes;
    if (siz>=20480) and (siz<100480) then s:= Format('%.1n '+sKBytes, [siz/1048]);
    if (siz>=100480) and (siz<1048576) then s:= Format('%u '+sKBytes, [siz div 1048]);
    if siz>=1048576  then s:= Format('%.1n '+SMBytes, [siz/1048576]);
    SGMails.Cells[4,i+1]:= s;
    // retrieve old selected message if still here to select it again
    if DisplayMails.GetItem(i).MessageUIDL= oldUIDL then SGMails.row:= i+1;
  end;

end;

procedure TFMailsInBox.MnuAccountPopup(Sender: TObject);
begin
  MnuMoveUp.Enabled:= not (LVAccounts.ItemIndex=0);
  if MnuAccountLog.Enabled then MnuAccountLog.ImageIndex:=0 else MnuAccountLog.ImageIndex:=1;
  if MnuGetAccMail.Enabled then MnuGetAccMail.ImageIndex:=2 else MnuGetAccMail.ImageIndex:=3;
  if MnuDeleteAcc.Enabled then MnuDeleteAcc.ImageIndex:=4 else MnuDeleteAcc.ImageIndex:=5 ;
  if MnuEditAcc.Enabled then MnuEditAcc.ImageIndex:=6 else MnuEditAcc.ImageIndex:=7;
  if MnuMoveUp.Enabled then MnuMoveUp.ImageIndex:=8 else MnuMoveUp.ImageIndex:=9 ;
  MnuMoveDown.Enabled:= not (LVAccounts.ItemIndex=LVAccounts.Items.count-1);
  if MnuMoveDown.Enabled then MnuMoveDown.ImageIndex:=10 else MnuMoveDown.ImageIndex:=11;
end;

procedure TFMailsInBox.MnuAnswerMsgClick(Sender: TObject);
var
  mndx: integer;
begin
  mndx:= SGMails.row-1;
  if mndx <0 then exit;
  OpenURL('mailto:'+DisplayMails.GetItem(mndx).MessageFrom+'<'+DisplayMails.GetItem(mndx).FromAddress+
          '>?subject=re:'+DisplayMails.GetItem(mndx).MessageSubject);
end;

procedure TFMailsInBox.MnuButtonBarPopup(Sender: TObject);
begin
  MnuDisplayMenu.Caption:= MMnuDisplayMenu.Caption;
  MnuDisplayBar.Caption:= MMnuDisplayBar.Caption;
end;

procedure TFMailsInBox.MnuDeleteMsgClick(Sender: TObject);
var
  andx: integer;
  mndx: integer;
  amndx: integer;
  uid2del: integer;
  CurAcc: TAccount;
  Subj: string;
begin
  mndx:= SGMails.row-1;
  if mndx<0 then exit;
  CurAcc:= FAccounts.Accounts.GetItemByUID(DisplayMails.GetItem(mndx).AccountUID);
  if CurAcc.Index<0 then exit
  else andx:= CurAcc.Index;
  Subj:= Copy(DisplayMails.GetItem(mndx).MessageSubject,1, 15)+'...';
  // Alarm before deleting
  if MsgDlg(Caption, Format(sAlertDelMmsg, [Subj]), mtWarning,
             [mbYes,mbNo], [YesBtn,NoBtn])= mrYes then
  begin
    // Add our new UID only if not already marked
    // Add UID to array
    amndx:= CurAcc.Mails.FindUIDL(DisplayMails.GetItem(mndx).MessageUIDL);
    if amndx>=0 then
    begin
      TAccount(FAccounts.Accounts.Items[andx]^).Mails.ModifyField(amndx, 'MessageToDelete', true);
      DisplayMails.ModifyField (mndx, 'MessageToDelete', true);
      uid2del:= length(CurAcc.UIDLToDel);
      Setlength(TAccount(FAccounts.Accounts.Items[andx]^).UIDLToDel, uid2del+1);
      TAccount(FAccounts.Accounts.Items[andx]^).UIDLToDel[uid2del]:= CurAcc.Mails.GetItem(amndx).MessageUIDL;
    end;
    PopulateMailsList(andx);
  end;
end;

procedure TFMailsInBox.MnuDisplayBarClick(Sender: TObject);
begin
  if TComponent(Sender)= MMnuDisplayBar then MnuDisplayBar.Checked:=  MMnuDisplayBar.Checked;
  if TComponent(Sender)= MnuDisplayBar then MMnuDisplayBar.Checked:= MnuDisplayBar.Checked;
  PnlToolbar.visible:= MnuDisplayBar.Checked;
  if not PnlToolbar.visible then
  begin
     FMailsInBox.Menu:= MainMnu;
     MnuDisplayMenu.checked:= true;
     MMnuDisplayMenu.Checked:= true;
     Fsettings.Settings.MenuBar:= true;
  end;
  Fsettings.Settings.ButtonBar:= PnlToolbar.visible;
end;

procedure TFMailsInBox.MnuDisplayMenuClick(Sender: TObject);
begin
  if TComponent(Sender)= MMnuDisplayMenu then MnuDisplayMenu.Checked:=  MMnuDisplayMenu.Checked;
  if TComponent(Sender)= MnuDisplayMenu then MMnuDisplayMenu.Checked:= MnuDisplayMenu.Checked;
  if MnuDisplayMenu.Checked then  FMailsInBox.Menu:= MainMnu else
  begin
    FMailsInBox.Menu:= nil;
    if not PnlToolbar.visible then
    begin
      PnlToolbar.visible:= true;
      Fsettings.Settings.ButtonBar:= true;
      MnuDisplayBar.Checked:= true;
      MMnuDisplayBar.Checked:= true;
    end;
  end;
  Fsettings.Settings.MenuBar:=  not (FMailsInBox.Menu= nil);
end;

procedure TFMailsInBox.MnuInfosClick(Sender: TObject);
var
  mndx: integer;
  mail:TMail;
  s: string;
begin
  mndx:= SGMails.row-1;
  if mndx<0 then exit;
  mail:= DisplayMails.GetItem(mndx);
  s:= SGMails.Columns[0].Title.Caption+': ';
  if length(Mail.MessageFrom)>0 then s:=s+Mail.MessageFrom+' ';
  s:=s+'('+mail.FromAddress+')'+LineEnding;
  s:=s+SGMails.Columns[1].Title.Caption+': '+Mail.AccountName+' ('+Mail.ToAddress+')'+LineEnding;
  s:=s+SGMails.Columns[2].Title.Caption+': '+Mail.MessageSubject+LineEnding;
  s:=s+SGMails.Columns[3].Title.Caption+': '+SGMails.Cells[3,SGMails.row]+LineEnding;
  s:=s+SGMails.Columns[4].Title.Caption+': '+SGMails.Cells[4,SGMails.row]+LineEnding;
  s:=s+'UID : '+Mail.MessageUIDL;
  ShowMessage(s);
end;

procedure TFMailsInBox.MnuMailsPopup(Sender: TObject);
var
  Subj: string;
begin
  if SGMails.row<1 then exit;
  Subj:= Copy(DisplayMails.GetItem(SGMails.row-1).MessageSubject,1, 15)+'...';
  MnuDeleteMsg.Caption:= Format(sMnuDelMsg, [Subj]);
  MnuAnswerMsg.Caption:= Format(sMnuAnswerMsg, [Subj]);
  if MnuInfos.Enabled then MnuInfos.ImageIndex:=0 else MnuInfos.ImageIndex:=1 ;
  if MnuAnswerMsg.Enabled then MnuAnswerMsg.ImageIndex:=2 else MnuAnswerMsg.ImageIndex:=3 ;
  if MnuDeleteMsg.Enabled then MnuDeleteMsg.ImageIndex:=4 else MnuMaximize.ImageIndex:=5;
end;

procedure TFMailsInBox.MnuMaximizeClick(Sender: TObject);
begin
  WindowState:=wsMaximized;
  Visible:= true;
end;

procedure TFMailsInBox.MnuIconizeClick(Sender: TObject);
begin
  Application.Minimize;
end;

procedure TFMailsInBox.MnuMoveClick(Sender: TObject);
var
  oldndx: integer;
  incr: integer;
begin
  oldndx:= LVAccounts.ItemIndex;
  if TMenuItem(Sender).Name='MnuMoveDown' then
  begin
    if (oldndx<0) or (oldndx>LVAccounts.Items.count-2) then exit;
    incr:= 1;
  end;
  if TMenuItem(Sender).Name='MnuMoveUp' then
  begin
    if (oldndx<=0) then exit;
    incr:= -1;
  end;
  FAccounts.Accounts.sorttype:= cdcNone;
  FAccounts.Accounts.ModifyField(oldndx, 'index', oldndx+incr);
  FAccounts.Accounts.ModifyField(oldndx+incr, 'index', oldndx);
  FAccounts.Accounts.sorttype:= cdcIndex;
  PopulateAccountsList(false);
  LVAccounts.ItemIndex:= oldndx+incr;
end;

procedure TFMailsInBox.MnuQuitClick(Sender: TObject);
begin
  CanClose:= true;
  close;
end;



procedure TFMailsInBox.MnuRestoreClick(Sender: TObject);
begin
  iconized:= false;
  visible:= true;
  WindowState:=wsNormal;
 //Need to reload position as it can change during hide in taskbar process
  left:= PrevLeft;
  top:= PrevTop;
  // Infos box is black when restore from tray (bug of TRichMemo ?)
  UpdateInfos;
  RMInfos.Invalidate;
  Application.BringToFront;
end;

procedure TFMailsInBox.MnuTrayPopup(Sender: TObject);
var
 imgndx:integer;
begin
  imgndx:=12;
  MnuRestore.Enabled:= (WindowState=wsMaximized) or (WindowState=wsMinimized);
  if MnuRestore.Enabled then MnuRestore.ImageIndex:=0 else MnuRestore.ImageIndex:=1 ;
  MnuMaximize.Enabled:= not (WindowState=wsMaximized);
  if MnuMaximize.Enabled then MnuMaximize.ImageIndex:=2 else MnuMaximize.ImageIndex:=3;
  MnuIconize.Enabled:= not (WindowState=wsMinimized);
  if MnuIconize.Enabled then MnuIconize.ImageIndex:= 4 else MnuIconize.ImageIndex:=5;
  if MnuGetAllMail.Enabled then MnuGetAllMail.ImageIndex:=6 else MnuGetAllMail.ImageIndex:=7;
  if MnuAbout.Enabled then MnuAbout.ImageIndex:=8 else MnuAbout.ImageIndex:=9;
  if MnuQuit.Enabled then MnuQuit.ImageIndex:=10 else MnuQuit.ImageIndex:=11;
  // Change images for mail client
  if AnsiContainsText(FSettings.Settings.MailClientName, 'outlook') then imgndx:= 14;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'thunder') then imgndx:= 16;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'gmail') then imgndx:= 18;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'windows') then imgndx:= 20;
  if AnsiContainsText(FSettings.Settings.MailClientName, 'outlook.com') then imgndx:= 22;
  if MnuLaunchClient.Enabled then MnuLaunchClient.ImageIndex:=imgndx else MnuLaunchClient.ImageIndex:=imgndx+1;
  MnuLaunchClient.Caption:= BtnLaunchClient.Hint;
end;

procedure TFMailsInBox.SGMailsBeforeSelection(Sender: TObject; aCol,
  aRow: Integer);
begin
  SGMails.Invalidate;
end;

procedure TFMailsInBox.SGMailsClick(Sender: TObject);
begin

end;

procedure TFMailsInBox.SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
var
  R: TRect;
  bmp: Tbitmap;
  bmppos: integer;
  AccCol: TColor;
  CurAcc: TAccount;
begin
  if arow=0 then exit;
  // remove selection highlight if the row has not the focus
  if SGMails.IsCellSelected[acol, arow] and SGHasFocus then
  begin
   SGMails.Canvas.Brush.Color := clHighlight;
   SGMails.Canvas.FillRect (ARect);
   SGMails.Canvas.font.Color:= clHighlightText;
  end else
  begin
    SGMails.Canvas.Brush.Color := clWindow;
    SGMails.Canvas.FillRect (ARect);
    SGMails.Canvas.font.Color:= clDefault;  ;
  end;
  // Get curent CurAcc color
  CurAcc:= FAccounts.Accounts.GetItemByUID(DisplayMails.GetItem(aRow-1).AccountUID);
  if CurAcc.Index>= 0 then AccCol:= CurAcc.Color
  else AccCol:= $FFFFFF;
  // Add mail image
  Case acol of
    0: begin
         bmppos:= 0;
         R.Left:= ARect.Left+2;
         R.Top:= ARect.Top+2;
         R.Right:=R.Left+18;
         R.Bottom:=R.Top+16;
         Bmp:= Tbitmap.Create;
         if DisplayMails.GetItem(aRow-1).MessageNew then bmppos:= 1;
         if Pos ('multipart', DisplayMails.GetItem(aRow-1).MessageContentType) >0 then
           bmppos:= bmppos+2;
         if DisplayMails.GetItem(aRow-1).MessageToDelete then bmppos:= 4;
         CropBitmap(MailPictures.Bitmap, bmp, bmppos );
         SGMails.Canvas.StretchDraw(R, bmp);
         if Assigned(Bmp) then Bmp.free;                                        // prevent access violation
         SGMails.Canvas.TextOut(ARect.Left+22,ARect.Top+3, SGMails.Cells[aCol, aRow]);
      end;
    1: begin
        SGMails.Canvas.Brush.Color := AccCol;
        SGMails.Canvas.Ellipse(Arect.Left,Arect.Top+5,Arect.Left+11, Arect.Top+16 );
        if SGMails.IsCellSelected[acol, arow] and SGHasFocus then
        begin
          SGMails.Canvas.Brush.Color := clHighlight
          //SGMails.Canvas.FillRect (ARect);
          //SGMails.Canvas.font.Color:= clHighlightText;
        end else
        begin
          SGMails.Canvas.Brush.Color := clWindow;
          //SGMails.Canvas.FillRect (ARect.left+12, Arect.Top, Arect.Right, Arect.Bottom);
          //SGMails.Canvas.font.Color:= clDefault;  ;
        end;
        SGMails.Canvas.TextOut(ARect.Left+13,ARect.Top+3, SGMails.Cells[aCol, aRow]);
       end
    else SGMails.Canvas.TextOut(ARect.Left+2,ARect.Top+3, SGMails.Cells[aCol, aRow]);
  end;

end;

// 3 next Procedures to remove selection highlight when the list has not the focus

procedure TFMailsInBox.SGMailsEnter(Sender: TObject);
begin
   SGHasFocus:= true;
end;

procedure TFMailsInBox.SGMailsExit(Sender: TObject);
begin
  SGHasFocus:= false;
  //SGMails.Invalidate;
end;

procedure TFMailsInBox.SGMailsKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

end;

procedure TFMailsInBox.SGMailsMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  pf: TPoint;
  CurCol, CurRow: integer;
begin
  CurCol:=0;
  CurRow:=0;
  SGMails.MouseToCell(X, Y, CurCol, CurRow);
  if (CurRow<0) or (CurCol<0) or (SGMails.RowCount<2) then exit;
  if Button = TMouseButton.mbRight then
  begin
    if visible then SGMails.SetFocus;
    SGMails.MouseToCell(X, Y, CurCol, CurRow);
    pf := SGMails.ClientToScreen(Point(X, Y));
    if CurRow>0 then
    begin
      SGMails.Row:= (CurRow);
      // Do not use the grids PopupMenu property, it
      // prevents this event handler comletely.
      // Instead, activate the menu manually here.
      MnuMails.Popup(pf.X, pf.Y);
    end;
  end;
  if (Button=TMouseButton.mbLeft) then
  begin
    if CurRow>0 then exit;  //Only first row
    SortMails(CurCol);
    PopulateMailsList(LVAccounts.ItemIndex);
  end;
end;

procedure TFMailsInBox.SortMails(CurCol: integer);
var
  i: integer;
begin
  Case CurCol of
    0: FSettings.Settings.MailSortTyp:= cdcMessageFrom;                   // first col sort on sender
    1: FSettings.Settings.MailSortTyp:= cdcMessageTo;
    2: FSettings.Settings.MailSortTyp:= cdcMessageSubject;
    3: FSettings.Settings.MailSortTyp:= cdcMessageDate;
    4: FSettings.Settings.MailSortTyp:= cdcMessageSize;
  end;
  for i:= 0 to SGMails.ColCount-1 do SGMails.Columns[i].Title.ImageIndex:= 0;
  if FSettings.Settings.MailSortDir=sdAscend then
  begin
    FSettings.Settings.MailSortDir:= sdDescend;
    SGMails.Columns[CurCol].Title.ImageIndex:= 2 ;
  end else
  begin
    FSettings.Settings.MailSortDir:= sdAscend;// sort order
    SGMails.Columns[CurCol].Title.ImageIndex:= 1 ;
  end;
end;

// Animate tray icon during checking mail  (TFPTimer)

procedure TFMailsInBox.OnChkMailTimer(Sender: TObject);
begin
  if ChkMailImgCount=0 then exit;
  if CheckingMail then
  try
    if ChkMailTimerTick < ChkMailImgCount then ILChkMail.GetIcon(ChkMailTimerTick, TrayMail.Icon);               //18/4/2025
    //TrayMail.Icon.Assign(ChkMailIcons[ChkMailTimerTick]);           //15/4/2025
  finally
    if ChkMailTimerTick < ChkMailImgCount-1 then inc(ChkMailTimerTick)
    else ChkMailTimerTick:=0;
  end;
end;

// Timer for time display (TFPTimer)

procedure TFMailsInBox.OnTimeTimer(Sender: TObject);
var
  s: string;
begin
  // DSisplay date and time every second
  s:= DateTimetoStr(now, StatusFmtSets);
  s[1]:= upCase(s[1]);        // First letter of day in uppercase
  LNow.Caption:= s;
end;

// ProcessMessage to avoid divide by zero (17/01/2025)

procedure TFMailsInBox.OnTrayTimer(Sender: TObject);
begin
  if not CheckingMail then
  begin
    if TrayTimerTick < ILTray.Count then
    begin
      ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
      Application.ProcessMessages;
      TrayMail.Icon.Assign(TrayTimerbmp);
    end;
    if TrayTimerTick < ILtray.count-1 then inc (TrayTimerTick, 1) else TrayTimerTick:= 0;
  end;
end;

procedure TFMailsInBox.BtnSettingsClick(Sender: TObject);
var
  i, oldlng, oldmailsel: integer;
  ndx: integer;
  smalltype: TBtnSize;
begin
  ndx:= LVAccounts.ItemIndex;
  smalltype:= bzNone;
  with FSettings do
  begin
    oldmailsel:=-1;
    CBStartup.Checked:= Settings.Startup;
    CBStartMini.Checked:= Settings.StartMini;
    CBSavSizePos.Checked:= Settings.SavSizePos;
    CBMailClientMini.Checked:= Settings.MailClientMini;
    CBHideInTaskBar.Checked:=Settings.HideInTaskbar;
    CBRestNewMsg.Checked:= Settings.RestNewMsg;
    CBSaveLogs.Checked:= Settings.SaveLogs;
    CBNoChkNewVer.Checked:= Settings.NoChkNewVer;
    CBStartupCheck.Checked:= Settings.StartupCheck;
    CBSmallBtns.Checked:= Settings.SmallBtns;
    CBNotifications.Checked:= Settings.Notifications;
    CBNoCloseAlert.checked:= Settings.NoCloseAlert;
    CBNoQuitAlert.Checked:= Settings.NoQuitAlert;
    CBDisplayAllAccMsgs.checked:= Settings.DisplayAllAccMsgs;

    GetDefaultMailCllient;    // and get mail clientrs list
    CBMailClient.Items.Clear; //Reset list of mail clients
    for i:=0 to length(MailClients)-1 do
    begin
      CBMailClient.Items.add(MailClients[i].Name);
      CBUrl.Checked:= MailClients[i].Url;
      if MailClients[i].tag=true  then
      begin
        oldmailsel:= i;   // selected mail client
        CBMailClient.ItemIndex:= oldmailsel;
      end;
      if (MailClients[i].Defaut=true) then
        if(oldmailsel<0) then CBMailClient.ItemIndex:= i; // default mail client
    end;
    ESoundFile.Text:= Settings.SoundFile;
    BtnPlaySound.Enabled:= not (length(ESoundFile.Text)=0);
    CBLangue.ItemIndex := LangNums.IndexOf(Settings.LangStr);
    oldlng := CBLangue.ItemIndex;
    if ShowModal= mrOK then
    begin
      Settings.Startup := CBStartup.Checked;
      Settings.StartMini := CBStartMini.Checked;
      Settings.SavSizePos := CBSavSizePos.Checked;
      Settings.MailClientMini:= CBMailClientMini.Checked;
      Settings.HideInTaskbar:= CBHideInTaskBar.Checked;
      Settings.RestNewMsg:= CBRestNewMsg.Checked;
      Settings.SaveLogs:= CBSaveLogs.Checked;
      Settings.NoChkNewVer := CBNoChkNewVer.Checked;
      Settings.StartupCheck:= CBStartupCheck.Checked;
      Settings.Notifications:= CBNotifications.Checked;
      Settings.NoCloseAlert:= CBNoCloseAlert.Checked;
      Settings.NoQuitAlert:= CBNoQuitAlert.Checked;
      Settings.DisplayAllAccMsgs:= CBDisplayAllAccMsgs.Checked;
      Settings.MailClient:= MailClients[CBMailClient.ItemIndex].Command;
      Settings.MailClientIsUrl:= CBUrl.Checked;
      Settings.SoundFile:= ESoundFile.Text;
      Settings.LangStr := LangNums.Strings[CBLangue.ItemIndex];
      // if one of them has changed, then need change buttons
      if (CBLangue.ItemIndex<>oldlng) or
         (Settings.MailClientName<>MailClients[CBMailClient.ItemIndex].Name) or
          (Settings.SmallBtns <> CBSmallBtns.Checked) then
      begin
        if (CBLangue.ItemIndex<>oldlng) then
        begin
          LangFile:= TBbIniFile.Create(ExtractFilePath(Application.ExeName) + 'lang'+PathDelim+Settings.LangStr+'.lng');
          self.Translate(LangFile);               // self is important !!! translate main form
          //ModLangue;
          GetMailClientNames(false);
          Application.QueueAsyncCall(@CheckUpdate, ChkVerInterval);
        end;
        if (Settings.MailClientName<>MailClients[CBMailClient.ItemIndex].Name) then
        begin
          Settings.MailClientName:= MailClients[CBMailClient.ItemIndex].Name;
        end;
        if Settings.SmallBtns <> CBSmallBtns.Checked then
        begin
          smalltype:= TBtnSize(integer(CBSmallBtns.Checked));
          Settings.SmallBtns:= CBSmallBtns.Checked;
        end;
        SetSmallBtns(smalltype);
      end;
      if SettingsChanged then
      begin
        PopulateAccountsList(false);  // Needed to change language on hints
        LVAccounts.ItemIndex:= ndx;
        LogAddLine(-1, now, sSettingsChange);
      end;
    end;
  end;
end;

procedure TFMailsInBox.ChkMailTimerStartTimer(Sender: TObject);
begin

end;

procedure TFMailsInBox.OnclickTimer(Sender: TObject);
begin
    // Clear double click watcher and timer
  doubleClick:= false;
  clickTimer.StopTimer;
  // Single click action
  BtnQuitClick(Sender);
end;


procedure TFMailsInBox.BtnAboutClick(Sender: TObject);
var
  chked: Boolean;
  alertmsg: String;
begin
  // If main windows is hidden, place the about box at the center of desktop,
  // else at the center of main windows
  if (Sender.ClassName= 'TMenuItem') and not visible then AboutBox.Position:= poDesktopCenter
  else AboutBox.Position:= poMainFormCenter;
  AboutBox.LastUpdate:= FSettings.Settings.LastUpdChk;
  chked:= AboutBox.Checked;
  AboutBox.ErrorMessage:='';
  if AboutBox.ShowModal= mrLast then
  begin
    UpdateDlg.sNewVer:= AboutBox.LastVersion;
    UpdateDlg.NewVersion:= true;
    {$IFDEF WINDOWS}
      if UpdateDlg.ShowModal = mryes then close;    // New version install experimental
    {$ELSE}
      OpenURL(AboutBox.UrlProgSite);
    {$ENDIF}
  end;
  FSettings.Settings.LastVersion:= AboutBox.LastVersion ;

  // If we have checked update and got an error
  if length(AboutBox.ErrorMessage)>0 then
  begin
    alertmsg := TranslateHttpErrorMsg(AboutBox.ErrorMessage, HttpErrMsgNames);
    if AlertDlg(Caption,  alertmsg, [OKBtn, CancelBtn, sNoLongerChkUpdates],
                    true, mtError)= mrYesToAll then FSettings.Settings.NoChkNewVer:= true;
    LogAddLine(-1, now, alertmsg);
  end;
  // Truncate date to avoid changes if there is the same day (hh:mm are in the decimal part of the date)
  if (not chked) and AboutBox.Checked then FSettings.Settings.LastVersion:= AboutBox.LastVersion;
  if trunc(AboutBox.LastUpdate) > trunc(FSettings.Settings.LastUpdChk) then
  begin
    FSettings.Settings.LastUpdChk:= AboutBox.LastUpdate;
    LogAddLine(-1, FSettings.Settings.LastUpdChk, AboutBox.LastVersion);
  end;
end;

procedure TFMailsInBox.BtnDeleteAccClick(Sender: TObject);
var
  ndx: Integer;
  AccName: String;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then
  begin
    ShowMessage(Format(sPlsSelectAcc, [sToDeleteAcc]));
    exit;
  end;
  AccName:= FAccounts.Accounts.GetItem(ndx).name;
  if AlertDlg (Caption, Format(sDeleteAccount, [AccName]), [OKBtn, CancelBtn, sAlertBoxCBNoShowAlert], false, mtWarning)= mrOK then
  begin
    FAccounts.Accounts.Delete(ndx);
    LogAddLine(-1, now, Format(sAccDeleted, [AccName]));
    Application.ProcessMessages;
    PopulateAccountsList(false);
  end;

end;

procedure TFMailsInBox.BtnGetAllMailClick(Sender: TObject);
var
  i: integer;
  ndx: Integer;
  CurAcc: TAccount;
begin
  if FAccounts.Accounts.count = 0 then exit;
  LogAddLine(-1, now, sCheckingAllMail);
  ndx:= LVAccounts.ItemIndex;   // Current selected CurAcc
  if ndx<0 then exit;
  MailChecking(true);
  Application.ProcessMessages;
  for i:= 0 to FAccounts.Accounts.count-1 do
  begin
    CurAccPend:= i;
    CurAcc:= FAccounts.Accounts.GetItem(i);
    if CurAcc.Enabled then
    begin
      GetPendingMail(i);
      Application.ProcessMessages;       //5/3/25
      if i=ndx then PopulateMailsList(i);
    end;
  end;
  MailChecking(false);
  Application.ProcessMessages;          // 5/3/25
  PopulateAccountsList(true);
  LVAccounts.ItemIndex:=ndx;
  if visible then LVAccounts.SetFocus;
end;

procedure TFMailsInBox.BtnHelpClick(Sender: TObject);
begin
  OpenDocument(HelpFile);
end;


procedure TFMailsInBox.BtnGetAccMailClick(Sender: TObject);
var
  ndx: integer;
begin
   ndx:= LVAccounts.ItemIndex;   // Current selected CurAcc
   if ndx>= 0 then GetAccMail(ndx);
end;

procedure TFMailsInBox.GetAccMail(ndx: integer);
begin
  if (ndx>=0) and not CheckingMail then
  begin
    MailChecking(true);
    CurAccPend:= ndx;
    Application.ProcessMessages;
    GetPendingMail(ndx);
    PopulateMailsList(ndx);
    MailChecking(false);
    PopulateAccountsList(true);
    LVAccounts.ItemIndex:=ndx;
    if visible then LVAccounts.SetFocus;
  end;
end;

// During mail checking prevent conflicts

function TFMailsInBox.MailChecking(status: boolean): boolean;
var
  ndx: integer;
begin
  result:= false;
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then exit;
  CheckingMail:= status;
  EnableControls(not status);
  if status then
  begin
    ChkMailTimerTick:= 0;
    ChkMailTimer.StartTimer;
    Application.ProcessMessages;
  end else
  begin
    ChkMailTimer.StopTimer;
    Application.ProcessMessages;
    TrayTimerTick:= 0;
    ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
    TrayMail.Hint:= sTrayHintNoMsg;
  end;
  result:= status;
end;

function TFMailsInBox.SetError(E: Exception; ErrorStr: String; ErrorUID: Integer; ErrorCaption: String; var ErrorsStr: String): boolean;
var
  ErrStr: String;
begin
  try
    ErrStr:= Format(ErrorStr, [E.Message]);
  except
    ErrStr:= Format(ErrorStr, ['Unknown error']);
  end;
  LStatus.Caption:= ErrorCaption+': '+ErrStr;
  LogAddLine(ErrorUID, now, LStatus.Caption);
  if ErrorsStr='' then ErrorsStr:= ErrStr
  else ErrorsStr:= ErrorsStr+LineEnding+ErrStr;
  result:= true;
end;

// retreive pop and imap mail

function TFMailsInBox.GetPendingMail(index: integer): Integer;
var
  msgs : Integer;
  i, j: integer;
  CurName: string;
  mails: TMailsList;
  min: TTime;
  CurAcc: TAccount;
  idMsgList: TIdMessageCollection;
  Err: boolean;
  ErrorsStr: string;
  slUIDL: TStringList;
  maildeleted: boolean;
begin
  result:= 0;
  msgs:= 0;
  if index<0 then exit;
  // reset error flag
  Err:= false;
  ErrorsStr:='';
  slUIDL:= TStringList.Create;
  mails:= TMailsList.create;
  idMsgList:= TIdMessageCollection.create;
  CurAcc:= FAccounts.Accounts.GetItem(index);
  LogAddLine(CurAcc.UID, now, Format(sCheckingAccMail, [CurAcc.Name]));
  TrayMail.Hint:= Format(sCheckingAccMail, [CurAcc.Name]);
  CurName:= CurAcc.Name;
  // Init protocol properties
  SetProtocolProperties(CurAcc);
  try
    LStatus.Caption:= Format(sConnectToServer, [Curacc.Name, CurAcc.Server]);
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
    Application.ProcessMessages;
    // Connect to server
    msgs:= ConnectServer(CurAcc, ErrorsStr);
    if msgs=-1 then
    begin
      err:=true;
      msgs:= 0;
    end;
    // Now retreive messages headers and mails infos
    if msgs > 0 then
      for i:= 1 to msgs do
      begin
        err:= not GetHeader(CurAcc, i, Mails, ErrorsStr);
        if err then continue;
      end;
    // delete and disconnect
    for i:=0 to msgs-1 do
    begin
      // delete messages with uidl in array, begins at message 1 and not 0 in pop3 protocol
      if length(CurAcc.UIDLToDel)>0 then
        for j:= 0 to length(CurAcc.UIDLToDel)-1 do
        begin
          maildeleted:= false;
          if CurAcc.Mails.GetItem(i).MessageUIDL= CurAcc.UIDLToDel[j]  then
          begin
            Case Curacc.Protocol of
              ptcPOP3: maildeleted:= IdPOP3_1.Delete(i+1);
              ptcIMAP: maildeleted:= IdIMAP4_1.UIDDeleteMsg(CurAcc.UIDLToDel[j]);
            end;
            if maildeleted then
            begin
              Mails.Delete(Mails.FindUIDL(CurAcc.UIDLToDel[j])) ;
              LogAddLine(CurAcc.UID, now, Format(sMsgDeleted, [CurAcc.Name, i+1]));
              TAccount(FAccounts.Accounts.Items[index]^).Mails.ModifyField(j, 'MessageToDelete',false);
              TAccount(FAccounts.Accounts.Items[index]^).UIDLToDel[j]:= '';
            end else LogAddLine(CurAcc.UID, now, Format(sMsgNotDeleted, [CurAcc.Name, i+1]));
          end;
        end;
    end;
    LStatus.Caption:= Format(sDisconnectServer, [Curacc.Name, CurAcc.Server]);
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
    Case Curacc.Protocol of
      ptcPOP3: IdPop3_1.Disconnect;
      ptcIMAP: IdIMAP4_1.Disconnect(true);
    end;
  except
    on E: Exception do
    begin
      err:= SetError(E, sConnectErrorMsg, CurAcc.UID, CurAcc.Name, ErrorsStr);
    end;
  end;
  Setlength(TAccount(FAccounts.Accounts.Items[index]^).UIDLToDel,0);
  if msgs>1 then begin
    LStatus.Caption:= CurName+' : '+Format(sMsgsFound, [msgs]);
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
  end else
  begin
    LStatus.Caption:= CurName+' : '+Format(sMsgFound, [msgs]) ;
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
  end;
  // Update CurAcc checkmail dates
  Setfire(curacc, now, fmLast);
  Application.ProcessMessages;
  min:= EncodeTime(0,FAccounts.Accounts.GetItem(index).interval,0,0);
  SetFire(CurAcc, now+min, fmNext);
  TAccount(FAccounts.Accounts.Items[index]^).Error:= Err;          // do not fire event
  TAccount(FAccounts.Accounts.Items[index]^).ErrorStr:= ErrorsStr; // do not fire event
  TAccount(FAccounts.Accounts.Items[index]^).Mails.Reset;          // do not fire event
  if Mails.count > 0 then
    for i:=0 to Mails.count-1 do
    begin
      TAccount(FAccounts.Accounts.Items[index]^).Mails.AddMail(Mails.GetItem(i));
      Application.ProcessMessages;
    end;
  if assigned (Mails) then Mails.free;
  if assigned(idMsgList) then idMsgList.free;
  if assigned(slUIDL) then slUIDL.free;
  result:= msgs;
end;

procedure TFMailsInBox.SetProtocolProperties(CurAcc: TAccount);
begin
  Case Curacc.Protocol of
    ptcPOP3:
      begin
        //IdPOP3_1.KeepAlive;
        IdPOP3_1.Host:= CurAcc.Server;
        IdPOP3_1.Port:= CurAcc.Port;
        IdPOP3_1.Username:= CurAcc.UserName;
        IdPOP3_1.Password:= CurAcc.Password;
        // Authentication method
        if Curacc.SecureAuth then IdPOP3_1.AuthType:= patSASL
        else IdPOP3_1.AuthType:= patUserPass;
      end;
    ptcIMAP:
      begin
        IdIMAP4_1.Host:= CurAcc.Server;
        IdIMAP4_1.Port:= CurAcc.Port;
        IdIMAP4_1.Username:= CurAcc.UserName;
        IdIMAP4_1.Password:= CurAcc.Password;
        // Authentication method
        if Curacc.SecureAuth then IdIMAP4_1.AuthType:= iatSASL
        else IdIMAP4_1.AuthType:= iatUserPass;
      end;
  end;
end;

// Returns -1 on error and update Errorstr

function TFMailsInBox.ConnectServer(CurAcc: TAccount; var ErrorsStr: string): integer;
var
  AMailBoxList: TStringList;
begin
  result:= -1;
  Case Curacc.Protocol of
    ptcPOP3:
      begin
        try
          if CurAcc.SSL>0 then
            IdPop3_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(idPop3_1);
           IdPop3_1.UseTLS := TIdUseTLS(CurAcc.SSL);
           //IdPOP3_1.ConnectTimeout:=;
           IdPOP3_1.Connect;
          //Application.ProcessMessages;
          result := IdPop3_1.CheckMessages;
         except
          on E: Exception do
          begin
            SetError(E, sConnectErrorMsg, CurAcc.UID, CurAcc.Name, ErrorsStr);
          end;
        end;
     end;
    ptcIMAP:
      begin
        AMailBoxList:= TStringList.create;
        try
          if CurAcc.SSL>0 then IdIMAP4_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(IdIMAP4_1);
          IdIMAP4_1.UseTLS := TIdUseTLS(CurAcc.SSL);
          if IdIMAP4_1.Connect then
          IdIMAP4_1.ListSubscribedMailBoxes(AMailBoxList);
          //Application.ProcessMessages;
          if not IdIMAP4_1.SelectMailBox('Inbox') then      // select proper mailbox
          begin
            // todo select another mailbox
          end;
          result:= IdIMAP4_1.MailBox.TotalMsgs;
        except
          on E: Exception do
          begin
            SetError(E, sConnectErrorMsg, CurAcc.UID, CurAcc.Name, ErrorsStr);
          end;
        end;
        if Assigned(AMailBoxList) then AMailBoxList.free;
      end;
  end;
end;

// Reurns false on error

function TFMailsInBox.GetHeader(CurAcc: TAccount; MailIndex: integer; var mails: TMailsList; var ErrorsStr: string): boolean;
var
  idMsg: TIdMessage;
  sUIDL: string;
  slUIDL: TStringList;
  siz: Integer;
  mail: Tmail;
begin
  result:= true;
  idMsg:= TIdMessage.Create(self);
  sUIDL:='';
  slUIDL:= TStringList.Create;
  mail:= Default(Tmail);
  Case Curacc.Protocol of
    ptcPOP3:
      begin
        try
          Result:= IdPop3_1.RetrieveHeader(MailIndex, idMsg);
          if Result then siz:= IdPop3_1.RetrieveMsgSize(MailIndex);
          IdPop3_1.UIDL(slUIDL, MailIndex);
          idMsg.UID:= slUIDL.Strings[0];
        except
          on E: Exception do
          begin
            result:= not SetError(E, sHeaderErrorMsg, CurAcc.UID, CurAcc.Name, ErrorsStr);
          end;
        end;
      end;
    ptcIMAP:
      begin
        try
          siz:= IdIMAP4_1.RetrieveMsgSize(MailIndex);
          IdIMAP4_1.RetrieveHeader(MailIndex, idMsg);

          IdIMAP4_1.GetUID(MailIndex, sUIDL);
          idMsg.UID:= sUIDL;
          siz:= siz+length(idMsg.Headers.Text) ;
          Application.ProcessMessages;
       except
          on E: Exception do
          begin
           result:= not SetError(E, sHeaderErrorMsg, CurAcc.UID, CurAcc.Name, ErrorsStr);
          end;
        end;
      end;
  end;
  Application.ProcessMessages;
  GetMailInfos(CurAcc, Mail, IdMsg, siz);
  if CurAcc.Mails.FindUIDL(Mail.MessageUIDL)>=0 then
    Mail.MessageNew:= false else Mail.MessageNew:= true;
  Mails.AddMail(Mail);
  if Assigned(idMsg) then idMsg.free;
  if Assigned(slUIDL) then slUIDL.free;
end;

// store lastfire and nextfire dates in the settings
// format : UID string '|' Unix time string

procedure TFMailsInBox.SetFire(Curacc: TAccount; datim: TDateTime; mode: TFireMode);
var
  uidfnd: integer;
  cnt: integer;
  i: integer;
  s, s1: string;
begin
  //if not CurAcc.Enabled then exit;
  uidfnd:= -1;
  s1:= InttoStr(CurAcc.UID)+'|'+ IntToStr(DateTimeToUnix(datim)); //TimeDateToString(datim, 'dd/mm/yyyy hh:nn:ss'); ;
  if mode= fmLast then cnt:= FSettings.Settings.LastFires.count
  else cnt:= FSettings.Settings.NextFires.count;
  if  cnt >0 then
   begin
     for i:= 0 to cnt-1 do
     begin
       if mode= fmLast then s:= FSettings.Settings.LastFires.Strings[i]
       else s:= FSettings.Settings.NextFires.Strings[i];
       if pos(InttoStr(CurAcc.UID), s) >0 then
       begin
         uidfnd:= i;
         break;
       end;
     end;
   end;
   if (mode= fmLast)then
   begin
     if uidfnd>=0 then FSettings.Settings.LastFires.Strings[uidfnd]:= s1
     else FSettings.Settings.LastFires.add(s1);
   end else
   begin
     if uidfnd>=0 then FSettings.Settings.NextFires.strings[uidfnd]:= s1
    else FSettings.Settings.NextFires.add(s1);
  end;
end;

// retrieve infos from mail header

procedure TFMailsInBox.GetMailInfos(CurAcc: TAccount; var Mail: TMail; IdMsg: TIdMessage; siz: Integer);
var
  sfrom: string;
begin
  try
  sfrom:= IdMsg.From.Name;
  if length(sfrom)=0 then sfrom:= idMsg.From.Address;
  Mail.AccountName:= CurAcc.Name ;
  Mail.AccountIndex:=CurAcc.Index;
  Mail.AccountUID:= CurAcc.UID;
  Mail.MessageFrom:= sfrom;
  Mail.FromAddress:= idMsg.From.Address;
  Mail.MessageUIDL:= idMsg.UID;
  Mail.MessageSubject:= idMsg.Subject;
  Mail.MessageTo:= idMsg.Recipients[0].Name;
  Mail.ToAddress:= idMsg.Recipients[0].Address;
  Mail.MessageDate:= idMsg.Date;
  Mail.MessageSize:= siz;
  Mail.MessageContentType:=  IdMsg.ContentType ;
  except
    on E: Exception do LStatus.Caption:= E.Message;
  end;
end;

// Import external accounts
// mailattente accounts
// Outlook 2007-2013 accounts (password is not retrieved)
// Thunderbird accounts (password is not retrieved)

procedure TFMailsInBox.BtnImportClick(Sender: TObject);
var
  i, j: integer;
  s: string;
begin
  with FImpex do
  begin
    if ShowModal=mrOK then
    begin
      j:=0;
      for i:=0 to ImpAccounts.count-1 do
      begin
        if LBImpex.Selected[i] then
        begin
          FAccounts.Accounts.AddAccount(ImpAccounts.GetItem(i));
          inc(j);
        end;
       end;
      FAccounts.Accounts.DoSort;
      PopulateAccountsList(false);
      if j>1 then s:= sAccountsImported else s:= sAccountImported;
      MsgDlg(Caption, Format(s, [j, CBAccType.Items[CBAccType.ItemIndex]]),
        mtInformation, [mbOK], [OKBtn], 0);
      LogAddLine(-1, now, Format(s, [j, CBAccType.Items[CBAccType.ItemIndex]]));
    end;
  end;
end;

procedure TFMailsInBox.BtnLaunchClientClick(Sender: TObject);
var
  exec: string;
  A: TStringArray;
  sl: TstringList;
  i: integer;
begin
  if Fsettings.Settings.MailClientMini then MnuIconizeClick(sender);
  // if it is an URL, open it in the default browser
  if (FSettings.Settings.MailClientIsUrl) then
  begin
    OpenUrl(FSettings.Settings.MailClient);
  end else     // it is an app, extract parameters and execute it
  begin
    A:= FSettings.Settings.MailClient.Split(' ','"');
    sl:= TstringList.Create;
    if length(A)>0 then
    begin
      exec:= A[0];
      if length(A)>1 then
      for i:=1 to length(A)-1 do
        sl.Add(A[I]);
      Execute(exec, sl);
    end;
    sl.free;
  end;
  if FSettings.Settings.MailClientMini then MnuIconizeClick(Sender);
end;

// Log display, all system log or only CurAcc log

procedure TFMailsInBox.BtnLogClick(Sender: TObject);
var
  Curacc:TAccount;
  LogDoc: TStringList;
  i: integer;
  s: string;
  sendername: string;
  LogText: string;
  // parse function filter : CurAcc uid
  function parseline(st:String; filter: string=''):String;
  var
    A:TStringArray;
  begin
    result:= '';
    A:= st.Split('|');
    if length(A)<3 then exit;   // must have 3 parts, avoid error if line is incomplete or empty
    if pos('**', A[2]) >0 then result:= A[2]+LineEnding      // separation line
    else
    begin
      if filter='' then
      begin
        result:= A[1]+' - '+A[2]+LineEnding;
      end else
      begin
        if A[0]=filter then result:= A[1]+' - '+A[2]+LineEnding else result:='';
      end;
    end;
  end;
  // End of parse function
begin
  // Concatenate with previous log if exists
  if FSettings.Settings.SaveLogs then LogText:= MainLog+SessionLog.text
  else LogText:= SessionLog.text;
  if Length(LogText)=0 then exit;
  LogDoc:= TstringList.Create;
  LogDoc.text:= LogText;
  s:='';
  sendername:= UpperCase(Tcomponent(sender).name);
  if (sendername= 'BTNACCOUNTLOG') or (sendername= 'MNUACCOUNTLOG') then
  begin
    if LVAccounts.ItemIndex>=0 then
    begin
       CurAcc:= FAccounts.Accounts.GetItem(LVAccounts.ItemIndex);
       for i:= 0 to logdoc.Count- 1 do s:=s+parseline(logdoc.strings[i], IntToStr(Curacc.UID));
    end else
    begin
      ShowMessage(Format(sPlsSelectAcc, [sToDisplayLog]));
      exit;
    end;
  end;
  if sendername= 'BTNLOG' then
  begin
    for i:= 0 to logdoc.Count -1  do s:=s+parseline(logdoc.strings[i]);
  end;
  With FLogView do
  begin
    if Copy(sendername,1,3)='BTN' then Caption:= TSpeedButton(Sender).Hint;
    if Copy(sendername,1,3)='MNU' then Caption:= TMenuItem(Sender).Caption;
    RMLog.rtf:='';
    RMLog.Text:=s;
    RMLog.SelStart:=0;
    RMLog.Sellength:=0;
    showmodal;
  end;
  if Assigned(LogDoc) then LogDoc.free;
end;

// Quit button click.
// Don't use with button event,
// fired from new buttonclick routine

procedure TFMailsInBox.BtnQuitClick(Sender: TObject);
begin
  if not FSettings.Settings.NoQuitAlert then
  begin
    Case AlertDlg(Caption, sNoQuitAlert, [OKBtn, CancelBtn, sAlertBoxCBNoShowAlert], true) of
       mrOk: begin
       end;
       mrYesToAll: begin
         FSettings.Settings.NoQuitAlert:= true;
         LogAddLine(-1, now, sNoShowQuitAlert);
       end;
    end;
  end ;
  MnuIconizeClick(Sender);

end;

// Quit button double click
// Don't use with button event,
// fired from new buttonclick routine

procedure TFMailsInBox.BtnQuitDblClick(Sender: TObject);
begin
  CanClose:= true;
  Close;
end;

// Procedures to differentiate button's single and double click
// based on mousedown event and clicktimer

procedure TFMailsInBox.BtnQuitMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if doubleClick then
  begin
    doubleClick:= false;
    timespan:= GetTickCount64-lastclick;
    // If double click is valid, respond
    if timespan < doubleClickMaxTime then
    begin
      clickTimer.StopTimer;
      BtnQuitDblClick(Sender)
    end;
    exit;
  end;
  // Double click was invalid, restart
  clickTimer.StopTimer;
  clickTimer.StartTimer;
  lastClick:= GetTickCount64;
  doubleClick:= true;
end;

// Disable controls during mail check to avoid conflicts
// Display hourglass cursor

procedure TFMailsInBox.EnableControls(Enable: boolean);
var
  i: integer;
begin
  if enable then
  begin
    PnlMails.Cursor:= crDefault;
    PnlAccounts.Cursor:= crDefault;
    for i:=0 to length(BtnsArr)-1 do BtnsArr[i].Btn.Enabled:=  BtnsArr[i].Enabled;
  end else
  begin
   PnlMails.Cursor:= crHourGlass;
   PnlAccounts.Cursor:= crHourGlass;
   for i:=0 to length(BtnsArr)-1 do
    begin
      BtnsArr[i].Enabled:= BtnsArr[i].Btn.Enabled;
      if (i=glHelp) or (i=glLaunchClient) then continue;
      if BtnsArr[i].Btn.Enabled then BtnsArr[i].Btn.Enabled:= enable;
    end;
  end;
  LVAccounts.Enabled:= enable;
  SGMails.Enabled:= enable;
  MnuAccountLog.Enabled:=BtnAccountLog.Enabled ;
  MnuLaunchClient.Enabled:= BtnLaunchClient.Enabled;
  MnuGetAllMail.Enabled:= BtnGetAllMail.Enabled;
  MnuDeleteMsg.Enabled:= Enable;
  MnuGetAccMail.Enabled:= BtnGetAccMail.Enabled;
  MnuDeleteAcc.Enabled:= BtnDeleteAcc.Enabled;
  MMnuFile.enabled:= enable;
  MMnuImport.Enabled:= BtnImport.Enabled;
  MMnuLaunchClient.Enabled:= BtnLaunchClient.Enabled;
  MMmnuQuit.Enabled:= BtnQuit.Enabled;
  MMnuMails.Enabled:= enable;
  MMnuGetAllMails.Enabled:= BtnGetAllMail.Enabled;
  MMnuGetAccMails.Enabled:= BtnGetAccMail.Enabled;
  MMnuDisplay.Enabled:= enable;
  MMnuDisplayMenu.Enabled:= enable;
  MMnuDisplayBar.Enabled:= enable;
  MMnuPrefs.Enabled:= enable;
  MMnuSettings.Enabled:= BtnSettings.Enabled;
  MMnuAddAcc.Enabled:= BtnAddAcc.Enabled;
  MMnuEditAcc.Enabled:= BtnEditAcc.Enabled;
  MMnuLog.Enabled:= BtnLog.Enabled;
  MMnuInfos.Enabled:= enable;
  MMnuHelp.Enabled:= BtnHelp.Enabled ;
  MMnuAbout.Enabled:= BtnAbout.Enabled;
  MnuDisplayMenu.Enabled:= enable;
  MnuDisplayBar.Enabled:= enable;

end;

// Draw a circle with messages count on the CurAcc and tray icon

procedure TFMailsInBox.DrawTheIcon(Bmp: Tbitmap; NewCount: integer; CircleColor: TColor);
var
  i : integer;
  s: string;
  tmpBmp: TBitmap;
begin
  // Bmp from Picture is 32bit, copy it to a temp bmp
  // we draw on the tmpBmp and finally assign it to input bmp
  tmpBmp:= Tbitmap.Create;
  tmpBmp.PixelFormat := pf24bit;
  tmpBmp.SetSize(Bmp.Width, Bmp.Height);
  tmpBmp.Canvas.Brush.Color := $FF00FF;
  tmpBmp.Canvas.FillRect(0, 0, tmpBmp.Width, tmpBmp.Height);
  tmpBmp.Canvas.Draw(0, 0, Bmp);
  With tmpBmp.Canvas do
  begin
    Brush.Style := bsSolid;
    Brush.Color := circlecolor;
    // Meme couleur que le masque
    If Brush.Color = $FF00FF then Brush.Color:= $FF01FF;
    Pen.Color := Brush.Color;
    Pen.Width:= 2;
    if NewCount > 9 then Rectangle(3,1,15,12)
    else Ellipse(0,1,12,13);
    // font
    Font.Name := 'Arial';
    Font.Style := [fsBold];
    // Clair ou foncé ?
    if DarkColor(Brush.color) then Font.Color := clWhite
    else     Font.Color := clBlack;
    //Font.Color := clWhite;
    Font.Size := 7;
    Brush.Style := bsClear;
    // number
    Case NewCount of
      -2: begin
            s:= '?';
            Font.Size := 8;
          end;
      -1: begin
            s:= '!';
            Font.Size := 8;
          end;
      else s:= IntToStr(NewCount);
    end;
    i := TextWidth(s) div 2;
    TextOut(6-i,1,s);
  end;
  Bmp.Assign(tmpBmp);
  if Assigned(tmpBmp) then tmpBmp.free;
end;


// Load control captions and text variable translations


procedure TFMailsInBox.Translate(LngFile: TBbInifile);
var
  prgName: String;
begin
  If Assigned(lngFile) then
  With LngFile do
  begin
    prgName:= ReadString('common', 'ProgName', 'Erreur');
    if prgName<>ProgName then ShowMessage(ReadString('common', 'ProgErr',
                         'Fichier de langue erroné. Réinstallez le programme'));
    OsVersion.Translate(LngFile);
     // general strings
    sRetConfBack:= ReadString('main','RetConfBack','Recharge la dernière configuration sauvegardée');
    sCreNewConf:= ReadString('main','CreNewConf','Création d''une nouvelle configuration');
    sLoadConf:= ReadString('main','LoadConf','Chargement de la configuration');
    OKBtn:= ReadString('common', 'OKBtn','OK');
    YesBtn:=ReadString('common','YesBtn','Oui');
    NoBtn:=ReadString('common','NoBtn','Non');
    CancelBtn:=ReadString('common','CancelBtn','Annuler');
    //Main Form  & components captions
    Caption:=ReadString('main','Caption','Courriels en attente');
    BtnImport.Hint:=ReadString('main','BtnImport.Hint',BtnImport.Hint );
    sBtnLogHint:=ReadString('main','BtnLogHint','Journal du compte %s');
    BtnLog.Hint:= ReadString('main', 'BtnLog.Hint', BtnLog.Hint);
    BtnGetAllMail.Hint:=ReadString('main','BtnGetAllMail.Hint',BtnGetAllMail.Hint);
    sBtnGetAccMailHint:=ReadString('main','BtnGetAccMailHint','Vérifier le compte %s');
    BtnLaunchClient.Hint:=ReadString('main','BtnLaunchClient.Hint', BtnLaunchClient.Hint);
    sBtnDeleteHint:=ReadString('main','BtnDeleteHint','Supprimer le compte %s');
    BtnAddAcc.Hint:=ReadString('main','BtnAddAcc.Hint',BtnAddAcc.Hint);
    sBtnEditAccHint:=ReadString('main','BtnEditAccHint','Modifier le compte %s');
    BtnSettings.Hint:=ReadString('main','BtnSettings.Hint',BtnSettings.Hint);
    BtnAbout.Hint:=ReadString('main','BtnAbout.Hint',BtnAbout.Hint);
    BtnHelp.Hint:=ReadString('main','BtnHelp.Hint',BtnHelp.Hint);
    BtnQuit.Hint:=ReadString('main','BtnQuit.Hint',BtnQuit.Hint);
    SGMails.Columns[0].Title.Caption:=ReadString('main','SGMails.Columns_0.Title.Caption',SGMails.Columns[0].Title.Caption);
    SGMails.Columns[1].Title.Caption:=ReadString('main','SGMails.Columns_1.Title.Caption',SGMails.Columns[1].Title.Caption);
    SGMails.Columns[2].Title.Caption:=ReadString('main','SGMails.Columns_2.Title.Caption',SGMails.Columns[2].Title.Caption);
    SGMails.Columns[3].Title.Caption:=ReadString('main','SGMails.Columns_3.Title.Caption',SGMails.Columns[3].Title.Caption);
    SGMails.Columns[4].Title.Caption:=ReadString('main','SGMails.Columns_4.Title.Caption',SGMails.Columns[4].Title.Caption);
    MnuDeleteMsg.Caption:=ReadString('main','MnuDeleteMsg.Caption',MnuDeleteMsg.Caption);
    MnuAnswerMsg.Caption:=ReadString('main','MnuAnswerMsg.Caption',MnuAnswerMsg.Caption);
    MnuInfos.Caption:=ReadString('main','MnuInfos.Caption',MnuInfos.Caption);
    sEmailCaption:=ReadString('main','EmailCaption','Courriel: %s');
    sLastCheckCaption:=ReadString('main','LastCheckCaption','Dernière vérification: %s');
    sNextCheckCaption:=ReadString('main','NextCheckCaption','Prochaine vérification: %s');
    sAccountImported:=ReadString('main','AccountImported','%d compte %s importé');
    sAccountsImported:=ReadString('main','AccountsImported','%d comptes %s importés');
    sMsgFound:=ReadString('main','MsgFound','%d message trouvé');
    sMsgsFound:=ReadString('main','MsgsFound','%d messages trouvés');
    sLStatusCaption:=ReadString('main','LStatusCaption','%s sur le compte %s le %');
    sConnectToServer:=ReadString('main','ConnectToServer','%s : Connexion au serveur %s');
    sConnectedToServer:=ReadString('main','ConnectedToServer','%s : Connecté au serveur %s');
    sConnectErrorMsg:=ReadString('main','ConnectErrorMsg','Erreur de connexion : %s');
    sHeaderErrorMsg:=ReadString('main','HeaderErrorMsg','Erreur d''obtention de l''entête : %s');
    sDisconnectServer:=ReadString('main','DisconnectServer','%s : Déonnexion du serveur %s');
    sDisconnectedServer:=ReadString('main','DisconnectedServer','%s : Déconnecté du serveur %s');
    sAccountChanged:=ReadString('main','AccountChanged', 'Le compte %s a été modifié');
    sAccountAdded:=ReadString('main','AccountAdded', 'Le compte %s a été ajouté');
    sAccountDisabled:=ReadString('main','AccountDisabled','Compte %s désactivé');
    sAccountEnabled:=ReadString('main','AccountEnabled','Compte %s activé');
    sLoadingAccounts:=ReadString('main','LoadingAccounts','Chargement des comptes');
    sCheckingAccMail:=ReadString('main','CheckingAccMail','Vérification du compte %s');
    sCheckingAllMail:=ReadString('main','CheckingAllMail','Vérification de tous les comptes actifs');
    sTrayHintNoMsg:=ReadString('main','TrayHintNoMsg','Aucun courriel en attente');
    sTrayHintMsg:=ReadString('main','TrayHintMsg','%s : %u courriel %%s');
    sTrayHintMsgs:=ReadString('main','TrayHintMsgs','%s : %u courriels %%s');
    sTrayHintNewMsg:=ReadString('main','TrayHintNewMsg','(%u nouveau)');
    sTrayHintNewMsgs:= ReadString('main','TrayHintNewMsgs','(%u nouveaux)');;
    TrayMail.BalloonTitle:=ReadString('main','TrayMail.BalloonTitle',TrayMail.BalloonTitle);
    sTrayBallHintMsg:=ReadString('main','TrayBallHintMsg','%s : %u nouveau courriel%s');
    sTrayBallHintMsgs:=ReadString('main','TrayBallHintMsgs','%s : %u nouveaux courriels%s');
    sNoShowCloseAlert:=ReadString('main','NoShowCloseAlert','Alerte de masquage de la fenêtre du programme déactivée');
    sNoShowQuitAlert:=ReadString('main','NoShowQuitAlert','Alerte de fermeture du programme désactivée');
    sNoQuitAlert:=ReadString('main','NoQuitAlert','Pour fermer le programme et ne plus vérifier l''arrivée '+
                                     'de nouveaux courriels, double cliquer sur ce bouton. Pour que la vérification '+
                                     'se poursuive en tâche de fond, cliquez sur le bouton '+
                                     '"Masquer la fenêtre de Courrier en attente".');
    sNoCloseAlert:=ReadString('main','NoCloseAlert','Le programme va se poursuivre en tâche de fond '+
                                     'pour vérifier l''arrivée de nouveaux courriels. Pour quitter le '+
                                     'programme, double cliquer sur le bouton "Quitter Courrier en attente".');
    sCannotQuit:=ReadString('main','CannotQuit','Impossible de quitter pendant la vérification de courriels');
    sClosingProg:=ReadString('main','ClosingProg','Fermeture de Courriels en attente');
    sRestart:=ReadString('main','Restart','Redémarrage après arrêt forcé');
    sSettingsChange:=ReadString('main','SettingsChange','Configuration modifiée');
    sMsgDeleted:=ReadString('main','MsgDeleted','%s : Message %u supprimé');
    sMsgNotDeleted:=ReadString('main','MsgNotDeleted','%s : Message %u non supprimé');
    sMnuDelMsg:=ReadString('main','MnuDelMsg', 'Effacer le courriel "%s"');
    sMnuAnswerMsg:=ReadString('main','MnuAnswerMsg','Répondre au courriel "%s"');
    sAlertDelMmsg:=ReadString('main','AlertDelMmsg','Voulez-vous supprimer le courriel "%s" ?');
    sAlertBoxCBNoShowAlert:=ReadString('main','AlertBoxCBNoShowAlert','Ne plus afficher cet avertissement');
    sBytes:=ReadString('main','Bytes','octets');
    sKBytes:=ReadString('main','KBytes','Ko');
    SMBytes:=ReadString('main','MBytes','Mo');
    sNeverChecked:=ReadString('main','NeverChecked','Pas encore vérifié');
    sUse64bit:=ReadString('main','Use64bit','Utilisez la version 64 bits de ce programme');
    sBtnLaunchClientDef:=ReadString('main','BtnLaunchClientDef', 'Lancer le client courrier');
    sBtnLaunchClientCust:=ReadString('main','BtnLaunchClientCust', 'Lancer %s');
    sDeleteAccount:=ReadString('main','DeleteAccount','Voulez-vous vraiment supprimer le compte %s ?');
    sCannotGetNewVerList:=ReadString('main','CannotGetNewVerList','Liste des nouvelles versions indisponible');
    sPlsSelectAcc:=ReadString('main','PlsSelectAcc','Sélectionnez un compte pour %s');
    sToDisplayLog:=ReadString('main','ToDisplayLog','pour pouvoir afficher son journal');
    sToDeleteAcc:=ReadString('main','ToDeleteAcc','pour pouvoir le supprimer');
    sAccDeleted:=ReadString('main','AccDeleted','Le compte %s a été supprimé');
    sToEditAcc:=ReadString('main','ToEditAcc','pour pouvoir le modifier');
    sNewAccount:=ReadString('main','NewAccount','Nouveau compte');
    sCreatedDataFolder:=ReadString('main','CreatedDataFolder','Dossier de données de Couriels en attente "%s" créé');
    // Main menu
    MMnuFile.Caption:=ReadString('main','MMnuFile.Caption',MMnuFile.Caption);
    MMnuMails.Caption:=ReadString('main','MMnuMails.Caption',MMnuMails.Caption);
    MMnuDisplay.Caption:=ReadString('main','MMnuDisplay.Caption',MMnuDisplay.Caption);
    MMnuDisplayMenu.Caption:=ReadString('main','MMnuDisplayMenu.Caption',MMnuDisplayMenu.Caption);
    MMnuDisplayBar.Caption:=ReadString('main','MMnuDisplayBar.Caption',MMnuDisplayBar.Caption);
    MMnuPrefs.Caption:=ReadString('main','MMnuPrefs.Caption', MMnuPrefs.Caption);
    MMnuInfos.Caption:=ReadString('main','MMnuInfos.Caption',MMnuInfos.Caption);
    HelpFile:= MIBExecPath+'help'+PathDelim+ReadString('main','HelpFile', 'mailsinbox.html');

    // About box
    AboutBox.LVersion.Hint:= OSVersion.VerDetail;
    AboutBox.Translate(LngFile);

    // UpdateDlg
    UpdateDlg.Translate (LangFile);

    // Alert
    sUpdateAlertBox:=ReadString('main','UpdateAlertBox','Version actuelle: %sUne nouvelle version %s est disponible. Cliquer pour la télécharger');
    sNoLongerChkUpdates:=ReadString('main','NoLongerChkUpdates','Ne plus rechercher les mises à jour');

    //Accounts
    FAccounts.Translate(LngFile);

    //Settings
    FSettings.Translate(LngFile);
    FSettings.Lstatus.Caption:= OSVersion.VerDetail;

    // Choose mail client
    FMailClientChoose.Translate(LngFile);

    // Impex
    FImpex.Translate(LngFile);

    // Logview
    FLogView.Translate(LngFile);

    // Tray
    MnuRestore.Caption:=ReadString('main','MnuRestore.Caption',MnuRestore.Caption);
    MnuMaximize.Caption:=ReadString('main','MnuMaximize.Caption',MnuMaximize.Caption);
    MnuIconize.Caption:=ReadString('main','MnuMinimize.Caption',MnuIconize.Caption);
    MnuGetAllMail.Caption:= BtnGetAllMail.Hint;
    MnuQuit.Caption:= ReadString('main','MnuQuit.Caption',MnuQuit.Caption);
    MnuAbout.Caption:=BtnAbout.Hint;

    // HTTP Error messages
    HttpErrMsgNames[0] := ReadString('HttpErr','SErrInvalidProtocol','Protocole "%s" invalide');
    HttpErrMsgNames[1] := ReadString('HttpErr','SErrReadingSocket','Erreur de lecture des données à partir du socket');
    HttpErrMsgNames[2] := ReadString('HttpErr','SErrInvalidProtocolVersion','Version de protocole invalide en réponse: %s');
    HttpErrMsgNames[3] := ReadString('HttpErr','SErrInvalidStatusCode','Code de statut de réponse invalide: %s');
    HttpErrMsgNames[4] := ReadString('HttpErr','SErrUnexpectedResponse','Code de statut de réponse non prévu: %s');
    HttpErrMsgNames[5] := ReadString('HttpErr','SErrChunkTooBig','Bloc trop grand');
    HttpErrMsgNames[6] := ReadString('HttpErr','SErrChunkLineEndMissing','Fin de ligne du bloc manquante');
    HttpErrMsgNames[7] := ReadString('HttpErr','SErrMaxRedirectsReached','Nombre maximum de redirections atteint: %s');
    // Socket error messages
    HttpErrMsgNames[8] := ReadString('HttpErr','strHostNotFound','Résolution du nom d''hôte pour "%s" impossible.');
    HttpErrMsgNames[9] := ReadString('HttpErr','strSocketCreationFailed','Echec de la création du socket: %s');
    HttpErrMsgNames[10] := ReadString('HttpErr','strSocketBindFailed','Echec de liaison du socket: %s');
    HttpErrMsgNames[11] := ReadString('HttpErr','strSocketListenFailed','Echec de l''écoute sur le port n° %s, erreur %s');
    HttpErrMsgNames[12]:=ReadString('HttpErr','strSocketConnectFailed','Echec de la connexion à %s');
    HttpErrMsgNames[13]:=ReadString('HttpErr','strSocketAcceptFailed','Connexion refusée d''un client sur le socket: %s, erreur %s');
    HttpErrMsgNames[14]:=ReadString('HttpErr','strSocketAcceptWouldBlock','La connexion pourrait bloquer le socket: %s');
    HttpErrMsgNames[15]:=ReadString('HttpErr','strSocketIOTimeOut','Impossible de fixer le timeout E/S à %s');
    HttpErrMsgNames[16]:=ReadString('HttpErr','strErrNoStream','Flux du socket non assigné');

  end;
end;

end.

