{******************************************************************************}
{ MailInBox main unit                                                          }
{ bb - sdtp - february 2020                                                     }
{ Check mails on pop3 and imap servers                                         }
{******************************************************************************}

unit mailsinbox1;

{$mode objfpc}{$H+}

interface

uses
  {$IFDEF WINDOWS}
  Win32Proc,
  {$ENDIF} Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, Grids, ComCtrls, Buttons, Menus, IdPOP3, IdSSLOpenSSL, LCLIntf,
  IdExplicitTLSClientServerBase, IdMessage, IdIMAP4, accounts1, lazbbutils,
  lazbbinifiles, lazbbosversion, LazUTF8, settings1, lazbbautostart,
  lazbbabout, Impex1, mailclients1, uxtheme, Types, IdComponent, fptimer,
  RichMemo, variants, IdMessageCollection, UniqueInstance, log1, registry,
  dateutils, strutils, lazbbchknewver;

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
    SplitterV: TSplitter;
    SplitterH: TSplitter;
    SGMails: TStringGrid;
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
    procedure BtnSettingsClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure DoChangeBounds(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BtnEditAccClick(Sender: TObject);
    procedure BtnCloseClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GetMailTimerTimer(Sender: TObject);
    procedure Id_client_Connected(Sender: TObject);
    procedure Id_client_Disconnected(Sender: TObject);
    procedure IdPOP3_1Status(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure LVAccountsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure MMnuClick(Sender: TObject);
    procedure MMnuImportClick(Sender: TObject);
    procedure MMnuPrefsClick(Sender: TObject);
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
    OS, OSTarget, CRLF: string;
    CompileDateTime: TDateTime;
    MIBAppDataPath: string;
    ProgName: string;
    LangStr: string;
    LangFile: TBbIniFile;
    LangNums: TStringList;
    LangFound: boolean;
    SettingsChanged: boolean;
    AccountsChanged:Boolean;
    ConfigFileName, AccountsFileName: string;
    ChkMailTimerTick: integer;
    CanClose: boolean;
    BaseUpdateUrl, ChkVerURL, version: string;
    sEmailCaption, sLastCheckCaption, sNextCheckCaption : string;
    sNoLongerChkUpdates, sLastUpdateSearch, sUpdateAvailable, sUpdateAlertBox: string;
    OKBtn, YesBtn, NoBtn, CancelBtn: string;
    sBtnLogHint, sBtnGetAccMailHint, sBtnDeleteHint, sBtnEditAccHint: string;
    BtnsArr : array of TBtnProps;
    sAccountImported, sAccountsImported: string;
    CheckingMail: boolean;
    SGHasFocus: boolean;
    sMsgFound, sMsgsFound: string;
    ChkMailTimer: TFPTimer;
    TimeTimer: TFPTimer;
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
    DisplayMails: TMailsList;
    HttpErrMsgNames: array [0..16] of string;
    sCannotGetNewVerList: string;
    procedure Initialize;
    procedure LoadCfgFiles(SettsFilename, AccsFilename: string);
    procedure SettingsOnChange(Sender: TObject);
    procedure SettingsOnStateChange(Sender: TObject);
    procedure SettingsOnQuitAlertChange(Sender:Tobject);
    procedure AccountsOnChange(Sender: TObject);
    function SaveConfig(Typ: TSaveMode): boolean;
    procedure PopulateAccountsList(notify: boolean);
    procedure ModLangue;
    procedure SetSmallBtns(small: TBtnSize);
    function GetPendingMail(index: integer): integer;
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
    procedure CheckUpdate;
    function AccFromMail(rw: integer): Integer;
    procedure SortMails(CurCol: integer);
  public
    OsInfo: TOSInfo;           //used by Settings1
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


// Intercept minimize system system command to correct
// wrong window placement on restore from tray

procedure TFMailsInBox.OnAppMinimize(Sender: TObject);
begin
  if FSettings.Settings.HideInTaskbar then
  begin
    PrevLeft:=left;
    PrevTop:= top;
    WindowState:= wsMinimized;
    Iconized:= HideOnTaskbar;
  end;
end;

// Intercept end session : save all pending settings and load the program
// on next logon (Windows only for the moment. On linux, we use autostart
// and we delete auttostrart on next startup

procedure TFMailsInBox.OnQueryendSession(var Cancel: Boolean);
var
  {$IFDEF WINDOWS}
    reg: TRegistry;
    RunRegKeyVal, RunRegKeySz: string;
  {$ENDIF}
  caClose: TCloseAction;
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
      Application.ProcessMessages;
      caClose:= caFree;
      FormClose(self, caClose);
    {$ENDIF}
    {$IFDEF Linux}
       SetAutostart(ProgName, Application.exename);
    {$ENDIF}
  end;
end;

// TFMailsInBox : This is the main form of the program

procedure TFMailsInBox.FormCreate(Sender: TObject);
var
  s: string;
begin
  // Variables initialization
  CanClose:= false;
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
  Application.OnMinimize:=@OnAppMinimize;
  Application.OnQueryEndSession:= @OnQueryendSession;
  // Initialize check mail timer
  ChkMailTimer:= TFPTimer.Create(self);
  ChkMailTimer.Interval:= 100;
  ChkMailTimer.UseTimerThread:= true;
  ChkMailTimer.Enabled:= true;
  ChkMailTimer.OnTimer:= @OnChkMailTimer;
  ChkMailTimerTick:= 0;
  // Initialize time display timer
  TimeTimer:= TFPTimer.Create(self);
  TimeTimer.Interval:= 1000;
  TimeTimer.UseTimerThread:= true;
  TimeTimer.Enabled:= true;
  TimeTimer.OnTimer:= @OnTimeTimer;
  TimeTimer.StartTimer;
  TrayTimerTick:=0;
  TrayTimerBmp:= TBitmap.Create;
  // Flag needed to execute once some processes in Form activation
  LangNums := TStringList.Create;
  SessionLog:= TStringList.Create;
  slLastFires:= TstringList.Create;
  slNextFires:=TstringList.Create;
  DisplayMails:= TMailsList.Create;
  {$IFDEF Linux}
    OS := 'Linux';
    CRLF := #10;
    LangStr := GetEnvironmentVariable('LANG');
    x := pos('.', LangStr);
    LangStr := Copy(LangStr, 0, 2);
    wxbitsrun := 0;
    OSTarget:= '';
    UserAppsDataPath := GetUserDir;
    // Get mail client
  {$ENDIF}
  {$IFDEF WINDOWS}
    OS := 'Windows ';
    CRLF := #13#10;
    // get user data folder
    s := ExtractFilePath(ExcludeTrailingPathDelimiter(GetAppConfigDir(False)));
    if Ord(WindowsVersion) < 7 then
      UserAppsDataPath := s                     // NT to XP
    else
    UserAppsDataPath := ExtractFilePath(ExcludeTrailingPathDelimiter(s)) + 'Roaming'; // Vista to W10
    LazGetShortLanguageID(LangStr);
    {$IFDEF WIN32}
      OSTarget := '32 bits';
    {$ENDIF}
    {$IFDEF WIN64}
      OSTarget := '64 bits';
    {$ENDIF}
  {$ENDIF}
  GetSysInfo(OsInfo);
  version := GetVersionInfo.ProductVersion;

  // Chargement des chaînes de langue...
  LangFile := TBbIniFile.Create(ExtractFilePath(Application.ExeName) + LowerCase(ProgName)+'.lng');
  // Cannot call Modlang as components are not yet created, use default language
  sOpenProgram:=LangFile.ReadString(LangStr,'OpenProgram','Ouverture de Courrier en attente');
  LogAddLine(-1, now, sOpenProgram+' - Version '+Version+ ' (' + OS + OSTarget + ')');
  LogAddLine(-1, now, OsInfo.VerDetail);
  MIBAppDataPath := UserAppsDataPath + PathDelim + ProgName + PathDelim;
  if not DirectoryExists(MIBAppDataPath) then CreateDir(MIBAppDataPath);
  LogFileName:= MIBAppDataPath+ProgName+'.log';
end;

// Add a line to log file. quote char is '|' to preserve real double quotes in log
// when reading with csv reader

function TFMailsInBox.LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
begin
  // values separated by '|'
 result:= SessionLog.Add(Inttostr(acc)+'|'+TimeDateToString(dat)+'|'+evnt);
end;

// Form activation only needed once

procedure TFMailsInBox.FormActivate(Sender: TObject);
begin
 if not Initialized then
 begin
   InitButtons;
   Initialize;
   CheckUpdate;
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
     //
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

procedure TFMailsInBox.CheckUpdate;
var
  errmsg: string;
  sNewVer: string;
  CurVer, NewVer: int64;
  alertpos: TPosition;
begin
  //Dernière recherche il y a plus de 7 jours ?
   errmsg := '';
   if (Trunc(Now)>Trunc(FSettings.Settings.LastUpdChk+7)) and (not FSettings.Settings.NoChkNewVer) then
   begin
     FSettings.Settings.LastUpdChk := Trunc(Now);
     sNewVer:= GetLastVersion(ChkVerURL, 'mailsinbox', errmsg);
     if length(sNewVer)=0 then
     begin
       if length(errmsg)=0 then MsgDlg(Caption, sCannotGetNewVerList,
           mtError,  [mbOK], [OKBtn])
       else MsgDlg(Caption, TranslateHttpErrorMsg(errmsg, HttpErrMsgNames ),
           mtError,  [mbOK], [OKBtn]);
       exit;
     end;
     NewVer := VersionToInt(sNewVer);
     // Cannot get new version
     if NewVer < 0 then exit;
     CurVer := VersionToInt(version);
     if NewVer > CurVer then
     begin
       AboutBox.LUpdate.Caption := Format(sUpdateAvailable, [sNewVer]);
       if not visible then alertpos:= poDesktopCenter
       else alertpos:= poMainFormCenter;
       Case AlertDlg(Caption, Format(sUpdateAlertBox,[version+LineEnding, sNewVer]), [OKBtn, CancelBtn, sNoLongerChkUpdates], true, mtInformation, alertpos) of
         mrOK: OpenURL(Format(BaseUpdateURl, [version, FSettings.Settings.LangStr]));
         mrAll: begin
           FSettings.Settings.NoChkNewVer := true;
           OpenURL(Format(BaseUpdateURl, [version, FSettings.Settings.LangStr]));
         end;
       end;
     end;
   end;
end;

// Initializing stuff

procedure TFMailsInBox.Initialize;
var
  i: integer;
  defmailcli: string;
  IniFile: TBbIniFile;
  tmplog: TstringList;
  CurAcc: TAccount;
  FilNamWoExt: string;
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
  // General pictures
  LaunchPicture:= Tpicture.Create;
  // Now, main settings
  FSettings.Settings.AppName:= LowerCase(ProgName);
  FAccounts.Accounts.AppName := LowerCase(ProgName);
  ConfigFileName:= MIBAppDataPath+'settings.xml';
  AccountsFileName:= MIBAppDataPath+'accounts.xml';
  FSettings.Settings.LangStr:= LangStr;
  ModLangue;
  //   PnlToolbar.visible:= true;
  FSettings.Settings.ButtonBar:= true;
  if not FileExists(AccountsFileName) then
  begin
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
      SaveConfig(All);
      LogAddLine(-1, now, sCreNewConf);
      //BtnGetAccMail.Enabled:= false;
      //BtnGetAllMail.Enabled:= False;
    end;
  end;
  // Check inifile with URLs, if not present, then use default
  IniFile:= TBbInifile.Create('mailsinbox.ini');
  BaseUpdateUrl:= IniFile.ReadString('urls', 'BaseUpdateUrl',
    'https://www.sdtp.com/versions/version.php?program=mailsinbox&version=%s&language=%s');
  ChkVerURL := IniFile.ReadString('urls', 'ChkVerURL',
    'https://www.sdtp.com/versions/versions.csv');
  if Assigned(IniFile) then IniFile.free;

  LoadCfgFiles(ConfigFileName, AccountsFileName);
  Application.Title:=Caption;
  if (OSINfo.Architecture = 'x86_64') and (OsTarget = '32 bits') then
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

  // AboutBox.UrlUpdate, AboutBox.LUpdate.Caption and Aboutbox.Caption
  // are done in ModLangue procedure
  AboutBox.Width:= 300; // to have more place for the long product name
  AboutBox.Image1.Picture.LoadFromResourceName(HInstance, 'ABOUTIMG');
  AboutBox.LProductName.Caption := GetVersionInfo.FileDescription;
  AboutBox.LCopyright.Caption := GetVersionInfo.CompanyName + ' - ' + DateTimeToStr(CompileDateTime);
  AboutBox.LVersion.Caption := 'Version: ' + Version + ' (' + OS + OSTarget + ')';
  AboutBox.UrlWebsite := GetVersionInfo.Comments;
  AboutBox.LUpdate.Hint := sLastUpdateSearch + ': ' + DateToStr(FSettings.Settings.LastUpdChk);
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
  FSettings.Settings.OnQuitAlertChange:= @SettingsOnQuitAlertChange;
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
  FSettings.LStatus.Caption := OsInfo.VerDetail;

  // set small buttons only if asked
  if FSettings.Settings.SmallBtns then SetSmallBtns(bzSmall)
  else SetSmallBtns(bzNone); // customize only launchmail button
  Constraints.MinWidth:= BtnQuit.left+BtnQuit.width+10;
  if width < Constraints.MinWidth then width := Constraints.MinWidth;
  // Get default mail client and stores its name
  defmailcli:= FSettings.GetDefaultMailCllient;
  // Launch config dailog to set mail client and other stuff
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
  SettingsOnQuitAlertChange(self);
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

// Change BtnQuit behaviour: one click if alert enabled,
// double click if aleret disabled
// procedure to call on each NoQuitAlert setting change

procedure TFMailsInBox.SettingsOnQuitAlertChange(sender: TObject);
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
end;

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
  // Change time display label position when window width change
  if Sender= self then LNow.left:= Clientwidth-LNow.width-10;
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

//procedure TFMailsInBox.LoadCfgFile(filename: string);
procedure TFMailsInBox.LoadCfgFiles(SettsFilename, AccsFilename: string);
var
  winstate: TWindowState;
  i: integer;
begin
  with FSettings do
  begin
    LogAddLine(-1, now, sLoadConf);
    Settings.LoadXMLFile(SettsFilename);
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
      self.WindowState := WinState;
      if Winstate = wsMinimized then
      begin
        Application.Minimize;
      end;
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
    if settings.StartMini then Application.Minimize;
    // Détermination de la langue (si pas dans settings, langue par défaut)
    if Settings.LangStr = '' then Settings.LangStr := LangStr;
    LangFile.ReadSections(LangNums);
    if LangNums.Count > 1 then
    begin
      CBLangue.Clear;;
      for i := 0 to LangNums.Count - 1 do
      begin
        FSettings.CBLangue.Items.Add(LangFile.ReadString(LangNums.Strings[i], 'Language',
          'Aucune'));
        if LangNums.Strings[i] = Settings.LangStr then
        begin
          LangFound := True;
          FSettings.CBLangue.ItemIndex:= i;
        end;
      end;
    end;


    // Si la langue n'est pas traduite, alors on passe en Anglais
    if not LangFound then
    begin
      Settings.LangStr := 'en';
    end;
  end;
  Modlangue;
  LogAddLine(-1, now, sLoadingAccounts);
  FAccounts.Accounts.Reset;
  FAccounts.Accounts.LoadXMLfile(AccsFilename);
  SettingsChanged := false;
end;


procedure TFMailsInBox.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  s: string;
  i: integer;
  curcolsw: string;
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
      // check if columns width has changed
      curcolsw:='';
      For i:= 0 to 4 do  curcolsw:= curcolsw+IntToHex(self.SGMails.Columns [i].Width, 4);
      if (curcolsw <> sColumnswidth) then DoChangeBounds(self);
      if FSettings.Settings.Startup then SetAutostart(progname, Application.exename)
      else UnSetAutostart(progname);
      if AccountsChanged then SaveConfig(All)
      else if SettingsChanged then SaveConfig(Setting) ;
      LogAddLine(-1, now, sClosingProg);
      LogAddLine(-1, now, '************************');
      if FSettings.Settings.SaveLogs then
      begin
        s:= Mainlog+SessionLog.text;
        SessionLog.text:=s;
      end;
      SessionLog.SaveToFile(LogFileName);
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
            LogAddLine(-1, now, sNoShowCloseAlert);
            MnuIconizeClick(sender);
         end;
       end;
    end else MnuIconizeClick(sender); ;
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
    //if FileExists (ConfigFileName) then
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
  //oldBallHint: string;
  totalNewMsgs: integer;
Begin
  if FAccounts.Accounts.Count = 0 then
  begin
    BtnGetAllMail.Enabled:= false;
    BtnGetAccMail.Enabled:= false;
    exit;
  end;
  BtnGetAllMail.Enabled:= true;
  BtnGetAccMail.Enabled:= true;

  sTrayNewHint:='';
  sTrayBallHint:='';
  TrayMail.Hint:= '';
  //oldBallHint:= TrayMail.BalloonHint;
  TrayMail.BalloonHint:='';
  sLineEnd:='';
  totalNewMsgs:=0;
  AccBmp:= TBitmap.create;
  LVAccounts.Clear;
  if Assigned(LVAccounts.SmallImages) then LVAccounts.SmallImages.Clear;
  ILTray.Clear;
  TrayPicture.LoadFromResourceName(HInstance, 'MAIL16');
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
      DrawTheIcon(TrayPicture.Bitmap, CurAcc.Mails.count, CurAcc.Color  );
      ILTray.AddMasked(TrayPicture.Bitmap, $FF00FF);  // modified icon
    end;
    if CurAcc.Error then
    begin
      DrawTheIcon(AccBmp, -1, CurAcc.Color);
      DrawTheIcon(TrayPicture.Bitmap, -1, CurAcc.Color);
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

// Event fired by account change

procedure TFMailsInBox.AccountsOnChange(Sender: TObject);
begin
  AccountsChanged := True;
end;

// edit or add an account

procedure TFMailsInBox.BtnEditAccClick(Sender: TObject);
var
  Account: TAccount;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx <0 then
  begin
     ShowMessage(Format(sPlsSelectAcc, [sToEditAcc]));
    exit;
  end;
  with FAccounts do
  begin
    if (TSpeedButton(Sender).Name='BtnAddAcc') or (TMenuItem(Sender).Name='MMnuAddAcc') then
    begin
      Account:= Default(TAccount);
      Account.UID:=0;
      Caption:= BtnAddAcc.Hint;
    end;
    if ((TSpeedButton(Sender).Name='BtnEditAcc') or (TMenuItem(Sender).Name='MnuEditAcc')
         or (TMenuItem(Sender).Name='MMnuEditAcc')) and  (ndx>=0) then
    begin
      Account:= Accounts.GetItem(ndx);
      Caption:= BtnEditAcc.Hint;
    end;
    EName.Text:= Account.Name;
    EServer.Text:= Account.Server;
    CBProtocol.ItemIndex:=ord(Account.Protocol);
    EUserName.Text:= Account.UserName;
    EPassword.Text:= Account.Password;
    EEmail.Text:= Account.Email;
    CBSSL.ItemIndex:= Account.SSL;
    EPort.Text:= IntToStr(Account.Port);
    CBSecureAuth.Checked:= Account.SecureAuth;
    EInterval.Value:= Account.Interval;
    EReplyEmail.Text:= Account.ReplyEmail;
    CBEnabledAcc.Checked:= Account.Enabled;
    CBColorAcc.Selected:= Account.Color;
    ESoundFile.Text:= Account.SoundFile;
    BtnPlaySound.Enabled:=not (length(ESoundFile.Text)=0);
    if ShowModal=mrOK then
    begin
      Account.Name:= EName.Text;
      Account.Server:= EServer.Text;
      Account.protocol:= TProtocols(CBProtocol.ItemIndex);
      Account.UserName:= EUserName.Text;
      Account.Password:= EPassword.Text;
      Account.Email:= EEmail.Text;
      Account.SSL:= CBSSL.ItemIndex;
      Account.Port:= StringToInt(EPort.Text);
      Account.SecureAuth:= CBSecureAuth.Checked;
      Account.Interval:= EInterval.Value;
      Account.ReplyEmail:= EReplyEmail.Text;
      Account.Enabled:= CBEnabledAcc.Checked;
      Account.Color:= CBColorAcc.Selected;
      if (TSpeedButton(Sender).Name='BtnEditAcc') and  (ndx>=0) then
      Accounts.ModifyAccount(ndx, Account)
      else Accounts.AddAccount(Account);
      PopulateAccountsList (false);
      if (TSpeedButton(Sender).Name='BtnEditAcc') then
      begin
        LVAccounts.ItemIndex:= ndx;
        LogAddLine(Account.UID, now, Format(sAccountChanged, [Account.Name]));
      end else
      begin
        LVAccounts.ItemIndex:= LVAccounts.Items.count-1 ;
        LogAddLine(Accounts.GetItem(LVAccounts.ItemIndex).UID , now, Format(sAccountAdded, [Account.Name]));
      end;
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
  if Assigned(ChkMailTimer) then ChkMailTimer.Destroy;
  if Assigned(SessionLog) then SessionLog.free;
  if Assigned(AccountPictures) then AccountPictures.Free;
  if Assigned(TrayPicture) then TrayPicture.Free;
  if Assigned(MailPictures) then MailPictures.Free;
  if Assigned(LaunchPicture) then LaunchPicture.Free;
  if Assigned(slLastFires) then slLastFires.Free;
  if Assigned(slNextFires) then slNextFires.Free;
  if Assigned(DisplayMails) then DisplayMails.free;
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
    // current account is enabled and interval defined
    CurAcc:= FAccounts.Accounts.GetItem(i);
    //if CurAcc.Enabled and (CurAcc.Interval>0) and (now> CurAcc.NextFire) then
    if CurAcc.Enabled and (CurAcc.Interval>0) and (now> GetFire(CurAcc, fmNext)) then
    begin
      min:= EncodeTime(0,CurAcc.interval,0,0);
      GetAccMail(i);
      //FAccounts.Accounts.ModifyField(i, 'NEXTFIRE', now+min);
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


procedure TFMailsInBox.LVAccountsSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  AccName: string;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx >= 0 then
  begin
    BtnGetAccMail.Enabled:= FAccounts.Accounts.GetItem(ndx).enabled;
    MnuGetAccMail.Enabled:= BtnGetAccMail.Enabled;
    LVAccounts.PopupMenu:= MnuAccount;
    AccName:= FAccounts.Accounts.GetItem(ndx).Name;
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
  PopulateMailsList(ndx);
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



procedure TFMailsInBox.MMnuImportClick(Sender: TObject);
begin

end;

procedure TFMailsInBox.MMnuPrefsClick(Sender: TObject);
begin

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
                   //FAccounts.Accounts.GetItem(ndx).Name, slastfire]);
      //RMInfos.Lines.add(Format(sNextCheckCaption, [TimeDateToString(CurAcc.NextFire)]));
      RMInfos.Lines.add(Format(sNextCheckCaption, [TimeDateToString(GetFire(CurAcc, fmNext))]));
                   //FAccounts.Accounts.GetItem(ndx).NextFire)]));
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
begin
  SGMails.RowCount:=1;
  DisplayMails.Reset;
  // If we display only selected accout messages
  if not FSettings.Settings.DisplayAllAccMsgs then
  begin
    if index<0 then exit;
    CurAcc:= FAccounts.Accounts.GetItem(index);
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
  j:= DisplayMails.count+1;
  SGMails.RowCount:= j;
  for i:=0 to DisplayMails.Count-1 do
  begin
    SGMails.Cells[0,i+1]:= DisplayMails.GetItem(i).MessageFrom;
    SGMails.Cells[1,i+1]:= DisplayMails.GetItem(i).AccountName;
    SGMails.Cells[2,i+1]:= DisplayMails.GetItem(i).MessageSubject;
    SGMails.Cells[3,i+1]:= TimeDateToString(DisplayMails.GetItem(i).MessageDate);
    // Change unit with size value
    siz:= DisplayMails.GetItem(i).MessageSize;
    if siz<20480 then s:= InttoStr(siz)+' '+sBytes;
    if (siz>=20480) and (siz<100480) then s:= Format('%.1n '+sKBytes, [siz/1048]);
    if (siz>=100480) and (siz<1048576) then s:= Format('%u '+sKBytes, [siz div 1048]);
    if siz>=1048576  then s:= Format('%.1n '+SMBytes, [siz/1048576]);
    SGMails.Cells[4,i+1]:= s;
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
  //mail:TMail;
  uid2del: integer;
  CurAcc: TAccount;
  Subj: string;
begin
  //andx:= LVAccounts.ItemIndex;
  andx:= AccFromMail(SGMails.Row);
  if andx<0 then exit;
  mndx:= SGMails.row-1;
  if mndx<0 then exit;
  CurAcc:= FAccounts.Accounts.GetItem(andx);
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

// retrieve account index from row in mail grid
function TFMailsInBox.AccFromMail(rw: integer): Integer;
var
  AccUID: integer;
  ndx: integer;
begin
  result:= -1;
  AccUID:= DisplayMails.getItem(rw-1).AccountUID;
  for ndx:=0 to FAccounts.Accounts.Count-1 do
  begin
    if FAccounts.Accounts.GetItem(ndx).UID=AccUID then
    begin
      result:= ndx;
      break;
    end;
  end;
end;

procedure TFMailsInBox.MnuMailsPopup(Sender: TObject);
var
  Subj: string;
  ndx: integer;
begin
  ndx:= AccFromMail(SGMails.row);
  if ndx<0 then exit;
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
  AccNdx: integer;
  AccCol: TColor;
begin
  if arow=0 then exit;
  // remove selection highlight if the control has not the focus
  //if not SGHasFocus then
  if SGMails.IsCellSelected[acol, arow] then
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
  // Get curent account color
  AccNdx:=DisplayMails.GetItem(aRow-1).AccountIndex;
  if (AccNdx>=0) and (AccNdx<FAccounts.Accounts.Count)
  then AccCol:= FAccounts.Accounts.GetItem(AccNdx).Color
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
         bmp.free;
         SGMails.Canvas.TextOut(ARect.Left+22,ARect.Top+3, SGMails.Cells[aCol, aRow]);
      end;
    1: begin
        SGMails.Canvas.Brush.Color := AccCol;
        SGMails.Canvas.Ellipse(Arect.Left,Arect.Top+5,Arect.Left+11, Arect.Top+16 );
        if SGMails.IsCellSelected[acol, arow] then //SGHasFocus then
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
  if CheckingMail then
  begin;
    TrayMail.Icon.LoadFromResourceID(HINSTANCE, ChkMailTimerTick);
    inc (ChkMailTimerTick);
    if ChkMailTimerTick > 5 then ChkMailTimerTick:=0;
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

procedure TFMailsInBox.OnTrayTimer(Sender: TObject);

begin
  if not CheckingMail then
  begin
    ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
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
          ModLangue;
          GetMailClientNames(false);
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


procedure TFMailsInBox.BtnAboutClick(Sender: TObject);
begin
  // If main windows is hidden, place the about box at the center of desktop,
  // else at the center of main windows
  if (Sender.ClassName= 'TMenuItem') and not visible then AboutBox.Position:= poDesktopCenter
  else AboutBox.Position:= poMainFormCenter;
  AboutBox.LastUpdate:= FSettings.Settings.LastUpdChk;
  AboutBox.ShowModal;
  // Truncate date to avoid changes if there is the same day (hh:mm are in the decimal part of the date)
  if trunc(AboutBox.LastUpdate) > trunc(FSettings.Settings.LastUpdChk) then
  begin
    FSettings.Settings.LastUpdChk:= AboutBox.LastUpdate;
    LogAddLine(-1, FSettings.Settings.LastUpdChk, sLastUpdateSearch);
  end;
end;

procedure TFMailsInBox.BtnDeleteAccClick(Sender: TObject);
var
  ndx: Integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then
  begin
    ShowMessage(Format(sPlsSelectAcc, [sToDeleteAcc]));
    exit;
  end;
  if AlertDlg (Caption, Format(sDeleteAccount, [FAccounts.Accounts.GetItem(ndx).name]), [OKBtn, CancelBtn, sAlertBoxCBNoShowAlert], false, mtWarning)= mrOK then
 begin
   FAccounts.Accounts.Delete(ndx);
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
  ndx:= LVAccounts.ItemIndex;   // Current selected account
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
      if i=ndx then PopulateMailsList(i);
    end;
  end;
  MailChecking(false);
  PopulateAccountsList(true);
  LVAccounts.ItemIndex:=ndx;
  if visible then LVAccounts.SetFocus;
end;

procedure TFMailsInBox.BtnHelpClick(Sender: TObject);
begin
  OpenDocument('help'+PathDelim+ProgName+'.html');
end;


procedure TFMailsInBox.BtnGetAccMailClick(Sender: TObject);
var
  ndx: integer;
begin
   ndx:= LVAccounts.ItemIndex;   // Current selected account
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
;
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then exit;
  CheckingMail:= status;
  EnableControls(not status);
  if status then
  begin
    //Screen
   // SGMails.Cursor:= crHourGlass;
   // SGMails.BeginUpdate;
    ChkMailTimerTick:= 0;
    ChkMailTimer.StartTimer;
  end else
  begin
    ChkMailTimer.StopTimer;
    TrayTimerTick:= 0;
    ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
    TrayMail.Hint:= sTrayHintNoMsg;
    //.. Screen.Cursor:= DefCursor;
  end;
  result:= status;
end;

// retreive pop and imap mail

function TFMailsInBox.GetPendingMail(index: integer): Integer;
var
  msgs : Integer;
  idMsg: TIdMessage;
  i, j, siz: integer;
  CurName: string;
  mail: TMail;
  mails: TMailsList;
  min: TTime;
  HeaderOK: boolean;
  CurAcc: TAccount;
  idMsgList: TIdMessageCollection;
  AMailBoxList: TStringList;
  Err: boolean;
  ErrStr, ErrorsStr: string;
  sUIDL: string;
  slUIDL: TStringList;
begin
  result:= 0;
  msgs:= 0;
  if index<0 then exit;
  // reset error flag
  Err:= false;
  ErrStr:= '';
  ErrorsStr:='';
  sUIDL:='';
  slUIDL:= TStringList.Create;
  mails:= TMailsList.create;
  idMsgList:= TIdMessageCollection.create;
  AMailBoxList:= TStringList.Create;
  CurAcc:= FAccounts.Accounts.GetItem(index);
  LogAddLine(CurAcc.UID, now, Format(sCheckingAccMail, [CurAcc.Name]));
  TrayMail.Hint:= Format(sCheckingAccMail, [CurAcc.Name]);
  CurName:= CurAcc.Name;
  Case Curacc.Protocol of
    ptcPOP3:
      begin
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
  try
    LStatus.Caption:= Format(sConnectToServer, [Curacc.Name, CurAcc.Server]);
    LogAddLine(CurAcc.UID, now, LStatus.Caption );
    Application.ProcessMessages;
    idMsg:= TIdMessage.Create(self);
    idMsgList:= TIdMessageCollection.Create;
    Case Curacc.Protocol of
      ptcPOP3:
        begin
          IdPop3_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(idPop3_1);
          IdPop3_1.UseTLS := TIdUseTLS(CurAcc.SSL);
          try
            IdPOP3_1.Connect;
            Application.ProcessMessages;
            msgs := IdPop3_1.CheckMessages;
           except
            on E: Exception do
            begin
              ErrStr:= Format(sConnectErrorMsg, [E.Message]);
              LStatus.Caption:= CurAcc.Name+': '+ErrStr;
              LogAddLine(CurAcc.UID, now, LStatus.Caption);
              if ErrorsStr='' then ErrorsStr:= ErrStr
              else ErrorsStr:= ErrorsStr+#10+ErrStr;
              Err:= true;
            end;
          end;
       end;
      ptcIMAP:
        begin
          IdIMAP4_1.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create(IdIMAP4_1);
          IdIMAP4_1.UseTLS := TIdUseTLS(CurAcc.SSL);
          try
            if IdIMAP4_1.Connect then
            IdIMAP4_1.ListSubscribedMailBoxes(AMailBoxList);
            Application.ProcessMessages;
            if not IdIMAP4_1.SelectMailBox('Inbox') then      // select proper mailbox
            begin
              // todo select proper mailbox
            end;
            msgs:= IdIMAP4_1.MailBox.TotalMsgs;
          except
            on E: Exception do
            begin
              ErrStr:= Format(sConnectErrorMsg, [E.Message]);
              LStatus.Caption:= CurAcc.Name+': '+ErrStr;
              LogAddLine(CurAcc.UID, now, LStatus.Caption);
              if ErrorsStr='' then ErrorsStr:= ErrStr
              else ErrorsStr:= ErrorsStr+#10+ErrStr;
              Err:= true;
            end;
          end;
        end;
    end;
    if msgs > 0 then
    begin
      for i:= 1 to msgs do
      begin
        Mail:= Default(TMail);
        Case Curacc.Protocol of
          ptcPOP3:
            begin
              try
                HeaderOK:= IdPop3_1.RetrieveHeader(i, idMsg);
                if HeaderOK then siz:= IdPop3_1.RetrieveMsgSize(i) else exit;
                IdPop3_1.UIDL(slUIDL, i);
                idMsg.UID:= slUIDL.Strings[0];
              except
                on E: Exception do
                begin
                  ErrStr:= Format(sHeaderErrorMsg, [E.Message]);
                  LStatus.Caption:= CurAcc.Name+': '+ErrStr;
                  LogAddLine(CurAcc.UID, now, LStatus.Caption);
                  if ErrorsStr='' then ErrorsStr:= ErrStr
                  else ErrorsStr:= ErrorsStr+#10+ErrStr;
                  Err:= true;
                end;
              end;
            end;
          ptcIMAP:
            begin
              try
                siz:= IdIMAP4_1.RetrieveMsgSize(i);
                IdIMAP4_1.RetrieveHeader(i, idMsg);

                IdIMAP4_1.GetUID(i, sUIDL);
                idMsg.UID:= sUIDL;
                siz:= siz+length(idMsg.Headers.Text) ;
                Application.ProcessMessages;
             except
                on E: Exception do
                begin
                  ErrStr:= Format(sHeaderErrorMsg, [E.Message]);
                  LStatus.Caption:= CurAcc.Name+': '+ErrStr;
                  LogAddLine(CurAcc.UID, now, LStatus.Caption);
                  if ErrorsStr='' then ErrorsStr:= ErrStr
                  else ErrorsStr:= ErrorsStr+#10+ErrStr;
                  Err:= true;
                end;
              end;
            end;
        end;
        Application.ProcessMessages;
        GetMailInfos(CurAcc, Mail, IdMsg, siz);
        if CurAcc.Mails.FindUIDL(Mail.MessageUIDL)>=0 then
        Mail.MessageNew:= false else Mail.MessageNew:= true;
        Mails.AddMail(Mail);
      end;
    end;
    Case Curacc.Protocol of
      ptcPOP3:
      begin
        for i:=0 to msgs-1 do
        begin
          // delete messages with uidl in array, begins at message 1 and not 0
          if length(CurAcc.UIDLToDel)>0 then
          begin
            for j:= 0 to length(CurAcc.UIDLToDel)-1 do
            begin
              if CurAcc.Mails.GetItem(i).MessageUIDL= CurAcc.UIDLToDel[j]  then
                if IdPOP3_1.Delete(i+1) then
                begin
                  Mails.Delete(Mails.FindUIDL(CurAcc.UIDLToDel[j])) ;
                  LogAddLine(CurAcc.UID, now, Format(sMsgDeleted, [CurAcc.Name, i+1]));
                  TAccount(FAccounts.Accounts.Items[index]^).Mails.ModifyField(j, 'MessageToDelete',false);
                end else LogAddLine(CurAcc.UID, now, Format(sMsgNotDeleted, [CurAcc.Name, i+1]));
            end;
          end;
        end;
        Setlength(TAccount(FAccounts.Accounts.Items[index]^).UIDLToDel, 0);
        LStatus.Caption:= Format(sDisconnectServer, [Curacc.Name, CurAcc.Server]);
        LogAddLine(CurAcc.UID, now, LStatus.Caption );
        IdPop3_1.Disconnect;
      end;
      ptcIMAP:
      begin
        for i:=0 to msgs-1 do
        begin
          //Message to delete ?
          if length(CurAcc.UIDLToDel)>0 then            // delete messages with uidl in array
          begin
            for j:=0 to length(CurAcc.UIDLToDel)-1 do
            begin
              if CurAcc.Mails.GetItem(i).MessageUIDL= CurAcc.UIDLToDel[j]  then
              if IdIMAP4_1.UIDDeleteMsg(CurAcc.UIDLToDel[j]) then
              begin
                Mails.Delete(Mails.FindUIDL(CurAcc.UIDLToDel[j])) ;
                LogAddLine(CurAcc.UID, now, Format(sMsgDeleted, [CurAcc.Name, i+1]));
                TAccount(FAccounts.Accounts.Items[index]^).Mails.ModifyField(j, 'MessageToDelete',false);
              end else LogAddLine(CurAcc.UID, now, Format(sMsgNotDeleted, [CurAcc.Name, i+1]));
            end;
          end;
        end;
        Setlength(TAccount(FAccounts.Accounts.Items[index]^).UIDLToDel, 0);
        LStatus.Caption:= Format(sDisconnectServer, [Curacc.Name, CurAcc.Server]);
        LogAddLine(CurAcc.UID, now, LStatus.Caption );
        IdIMAP4_1.Disconnect(true);
      end;
    end;
  except
    on E: Exception do
    begin
      ErrStr:= Format(sConnectErrorMsg, [E.Message]);
      LStatus.Caption:= CurAcc.Name+': '+ErrStr;
      LogAddLine(CurAcc.UID, now, LStatus.Caption);
      if ErrorsStr='' then ErrorsStr:= ErrStr
      else ErrorsStr:= ErrorsStr+#10+ErrStr;
      err:= true;
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
  // Update account checkmail dates
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
      //if length(Mails.GetItem(i).MessageUIDL)> 0 then
        TAccount(FAccounts.Accounts.Items[index]^).Mails.AddMail(Mails.GetItem(i));
      Application.ProcessMessages;
    end;
  if assigned (Mails) then Mails.free;
  if assigned(idMsgList) then idMsgList.free;
  if assigned(slUIDL) then slUIDL.free;
  result:= msgs;
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



// Log display, all system log or only account log

procedure TFMailsInBox.BtnLogClick(Sender: TObject);
var
  Curacc:TAccount;
  LogDoc: TStringList;
  i: integer;
  s: string;
  sendername: string;
  LogText: string;
  // parse function filter : account uid
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

procedure TFMailsInBox.BtnQuitClick(Sender: TObject);
begin
  if not FSettings.Settings.NoQuitAlert then
  begin
    Case AlertDlg(Caption, sNoQuitAlert, [OKBtn, CancelBtn, sAlertBoxCBNoShowAlert], true) of
       mrOk: begin
         CanClose:= true;
         Close;
       end;
       mrYesToAll: begin
         FSettings.Settings.NoQuitAlert:= true;
         LogAddLine(-1, now, sNoShowQuitAlert);
         CanClose:= true;
         Close;
       end;
    end;
  end;
end;

// Quit button double click

procedure TFMailsInBox.BtnQuitDblClick(Sender: TObject);
begin
  CanClose:= true;
  Close;
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

// Draw a circle with messages count on the account and tray icon

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
// from mailsinbox.lng

procedure TFMailsInBox.ModLangue;
begin
  with LangFile do
  begin
    LangStr:=FSettings.Settings.LangStr;
    // general strings
    sRetConfBack:= ReadString(LangStr,'RetConfBack','Recharge la dernière configuration sauvegardée');
    sCreNewConf:= ReadString(LangStr,'CreNewConf','Création d''une nouvelle configuration');
    sLoadConf:= ReadString(LangStr,'LoadConf','Chargement de la configuration');
    OKBtn:= ReadString(LangStr, 'OKBtn','OK');
    YesBtn:=ReadString(LangStr,'YesBtn','Oui');
    NoBtn:=ReadString(LangStr,'NoBtn','Non');
    CancelBtn:=ReadString(LangStr,'CancelBtn','Annuler');

    //Main Form  & components captions
    Caption:=ReadString(LangStr,'Caption','Courriels en attente');
    BtnImport.Hint:=ReadString(LangStr,'BtnImport.Hint',BtnImport.Hint );
    sBtnLogHint:=ReadString(LangStr,'BtnLogHint','Journal du compte %s');
    BtnGetAllMail.Hint:=ReadString(LangStr,'BtnGetAllMail.Hint',BtnGetAllMail.Hint);
    sBtnGetAccMailHint:=ReadString(LangStr,'BtnGetAccMailHint','Vérifier le compte %s');
    BtnLaunchClient.Hint:=ReadString(LangStr,'BtnLaunchClient.Hint', BtnLaunchClient.Hint);
    sBtnDeleteHint:=ReadString(LangStr,'BtnDeleteHint','Supprimer le compte %s');
    BtnAddAcc.Hint:=ReadString(LangStr,'BtnAddAcc.Hint',BtnAddAcc.Hint);
    sBtnEditAccHint:=ReadString(LangStr,'BtnEditAccHint','Modifier le compte %s');
    BtnSettings.Hint:=ReadString(LangStr,'BtnSettings.Hint',BtnSettings.Hint);
    BtnAbout.Hint:=ReadString(LangStr,'BtnAbout.Hint',BtnAbout.Hint);
    BtnHelp.Hint:=ReadString(LangStr,'BtnHelp.Hint',BtnHelp.Hint);
    BtnQuit.Hint:=ReadString(LangStr,'BtnQuit.Hint',BtnQuit.Hint);
    SGMails.Columns[0].Title.Caption:=ReadString(LangStr,'SGMails.Columns_0.Title.Caption',SGMails.Columns[0].Title.Caption);
    SGMails.Columns[1].Title.Caption:=ReadString(LangStr,'SGMails.Columns_1.Title.Caption',SGMails.Columns[1].Title.Caption);
    SGMails.Columns[2].Title.Caption:=ReadString(LangStr,'SGMails.Columns_2.Title.Caption',SGMails.Columns[2].Title.Caption);
    SGMails.Columns[3].Title.Caption:=ReadString(LangStr,'SGMails.Columns_3.Title.Caption',SGMails.Columns[3].Title.Caption);
    SGMails.Columns[4].Title.Caption:=ReadString(LangStr,'SGMails.Columns_4.Title.Caption',SGMails.Columns[4].Title.Caption);
    MnuDeleteMsg.Caption:=ReadString(LangStr,'MnuDeleteMsg.Caption',MnuDeleteMsg.Caption);
    MnuAnswerMsg.Caption:=ReadString(LangStr,'MnuAnswerMsg.Caption',MnuAnswerMsg.Caption);
    MnuInfos.Caption:=ReadString(LangStr,'MnuInfos.Caption',MnuInfos.Caption);
    //sAccountCaption:=ReadString(LangStr,'AccountCaption','Compte: %s');
    sEmailCaption:=ReadString(LangStr,'EmailCaption','Courriel: %s');
    sLastCheckCaption:=ReadString(LangStr,'LastCheckCaption','Dernière vérification: %s');
    sNextCheckCaption:=ReadString(LangStr,'NextCheckCaption','Prochaine vérification: %s');
    sAccountImported:=ReadString(LangStr,'AccountImported','%d compte %s importé');
    sAccountsImported:=ReadString(LangStr,'AccountsImported','%d comptes %s importés');
    sMsgFound:=ReadString(LangStr,'MsgFound','%d message trouvé');
    sMsgsFound:=ReadString(LangStr,'MsgsFound','%d messages trouvés');
    sLStatusCaption:=ReadString(LangStr,'LStatusCaption','%s sur le compte %s le %');
    sConnectToServer:=ReadString(LangStr,'ConnectToServer','%s : Connexion au serveur %s');
    sConnectedToServer:=ReadString(LangStr,'ConnectedToServer','%s : Connecté au serveur %s');
    sConnectErrorMsg:=ReadString(LangStr,'ConnectErrorMsg','Erreur de connexion : %s');
    sHeaderErrorMsg:=ReadString(LangStr,'HeaderErrorMsg','Erreur d''obtention de l''entête : %s');
    sDisconnectServer:=ReadString(LangStr,'DisconnectServer','%s : Déonnexion du serveur %s');
    sDisconnectedServer:=ReadString(LangStr,'DisconnectedServer','%s : Déconnecté du serveur %s');
    sAccountChanged:=ReadString(LangStr,'AccountChanged', 'Le compte %s a été modifié');
    sAccountAdded:=ReadString(LangStr,'AccountAdded', 'Le compte %s a été ajouté');
    sAccountDisabled:=ReadString(LangStr,'AccountDisabled','Compte %s désactivé');
    sAccountEnabled:=ReadString(LangStr,'AccountEnabled','Compte %s activé');
    sLoadingAccounts:=ReadString(LangStr,'LoadingAccounts','Chargement des comptes');
    sCheckingAccMail:=ReadString(LangStr,'CheckingAccMail','Vérification du compte %s');
    sCheckingAllMail:=ReadString(LangStr,'CheckingAllMail','Vérification de tous les comptes actifs');
    sTrayHintNoMsg:=ReadString(LangStr,'TrayHintNoMsg','Aucun courriel en attente');
    sTrayHintMsg:=ReadString(LangStr,'TrayHintMsg','%s : %u courriel %%s');
    sTrayHintMsgs:=ReadString(LangStr,'TrayHintMsgs','%s : %u courriels %%s');
    sTrayHintNewMsg:=ReadString(LangStr,'TrayHintNewMsg','(%u nouveau)');
    sTrayHintNewMsgs:= ReadString(LangStr,'TrayHintNewMsgs','(%u nouveaux)');;
    TrayMail.BalloonTitle:=ReadString(LangStr,'TrayMail.BalloonTitle',TrayMail.BalloonTitle);
    sTrayBallHintMsg:=ReadString(LangStr,'TrayBallHintMsg','%s : %u nouveau courriel%s');
    sTrayBallHintMsgs:=ReadString(LangStr,'TrayBallHintMsgs','%s : %u nouveaux courriels%s');
    sNoShowCloseAlert:=ReadString(LangStr,'NoShowCloseAlert','Alerte de masquage de la fenêtre du programme déactivée');
    sNoShowQuitAlert:=ReadString(LangStr,'NoShowQuitAlert','Alerte de fermeture du programme désactivée');
    sNoQuitAlert:=ReadString(LangStr,'NoQuitAlert','Pour fermer le programme et ne plus vérifier l''arrivée '+
                                     'de nouveaux courriels, double cliquer sur ce bouton. Pour que la vérification '+
                                     'se poursuive en tâche de fond, cliquez sur le bouton '+
                                     '"Masquer la fenêtre de Courrier en attente".');
    sNoCloseAlert:=ReadString(LangStr,'NoCloseAlert','Le programme va se poursuivre en tâche de fond '+
                                     'pour vérifier l''arrivée de nouveaux courriels. Pour quitter le '+
                                     'programme, double cliquer sur le bouton "Quitter Courrier en attente".');
    sCannotQuit:=ReadString(LangStr,'CannotQuit','Impossible de quitter pendant la vérification de courriels');
    sClosingProg:=ReadString(LangStr,'ClosingProg','Fermeture de Courriels en attente');
    sRestart:=ReadString(LangStr,'Restart','Redémarrage après arrêt forcé');
    sSettingsChange:=ReadString(LangStr,'SettingsChange','Configuration modifiée');
    sMsgDeleted:=ReadString(LangStr,'MsgDeleted','%s : Message %u supprimé');
    sMsgNotDeleted:=ReadString(LangStr,'MsgNotDeleted','%s : Message %u non supprimé');
    sMnuDelMsg:=ReadString(LangStr,'MnuDelMsg', 'Effacer le courriel "%s"');
    sMnuAnswerMsg:=ReadString(LangStr,'MnuAnswerMsg','Répondre au courriel "%s"');
    sAlertDelMmsg:=ReadString(LangStr,'AlertDelMmsg','Voulez-vous supprimer le courriel "%s" ?');
    sAlertBoxCBNoShowAlert:=ReadString(LangStr,'AlertBoxCBNoShowAlert','Ne plus afficher cet avertissement');
    sBytes:=ReadString(LangStr,'Bytes','octets');
    sKBytes:=ReadString(LangStr,'KBytes','Ko');
    SMBytes:=ReadString(LangStr,'MBytes','Mo');
    sNeverChecked:=ReadString(LangStr,'NeverChecked','Pas encore vérifié');
    sUse64bit:=ReadString(LangStr,'Use64bit','Utilisez la version 64 bits de ce programme');
    sBtnLaunchClientDef:=ReadString(LangStr,'BtnLaunchClientDef', 'Lancer le client courrier');
    sBtnLaunchClientCust:=ReadString(LangStr,'BtnLaunchClientCust', 'Lancer %s');
    sDeleteAccount:=ReadString(LangStr,'DeleteAccount','Voulez-vous vraiment supprimer le compte %s ?');
    sCannotGetNewVerList:=ReadString(LangStr,'CannotGetNewVerList','Liste des nouvelles versions indisponible');
    sPlsSelectAcc:=ReadString(LangStr,'PlsSelectAcc','Sélectionnez un compte pour %s');
    sToDisplayLog:=ReadString(LangStr,'ToDisplayLog','pour pouvoir afficher son journal');
    sToDeleteAcc:=ReadString(LangStr,'ToDeleteAcc','pour pouvoir le supprimer');
    sToEditAcc:=ReadString(LangStr,'ToEditAcc','pour pouvoir le modifier');
    // Main menu
    MMnuFile.Caption:=ReadString(LangStr,'MMnuFile.Caption',MMnuFile.Caption);
    MMnuMails.Caption:=ReadString(LangStr,'MMnuMails.Caption',MMnuMails.Caption);
    MMnuDisplay.Caption:=ReadString(LangStr,'MMnuDisplay.Caption',MMnuDisplay.Caption);
    MMnuDisplayMenu.Caption:=ReadString(LangStr,'MMnuDisplayMenu.Caption',MMnuDisplayMenu.Caption);
    MMnuDisplayBar.Caption:=ReadString(LangStr,'MMnuDisplayBar.Caption',MMnuDisplayBar.Caption);
    MMnuPrefs.Caption:=ReadString(LangStr,'MMnuPrefs.Caption', MMnuPrefs.Caption);
    MMnuInfos.Caption:=ReadString(LangStr,'MMnuInfos.Caption',MMnuInfos.Caption);

    // About
    sNoLongerChkUpdates:=ReadString(LangStr,'NoLongerChkUpdates','Ne plus rechercher les mises à jour');
    sLastUpdateSearch:=ReadString(LangStr,'LastUpdateSearch','Dernière recherche de mise à jour');
    sUpdateAvailable:=ReadString(LangStr,'UpdateAvailable','Nouvelle version %s disponible');
    sUpdateAlertBox:=ReadString(LangStr,'UpdateAlertBox','Version actuelle: %sUne nouvelle version %s est disponible');
    Aboutbox.Caption:=ReadString(LangStr,'Aboutbox.Caption','A propos de Courriels en attente');
    AboutBox.LUpdate.Caption:=ReadString(LangStr,'AboutBox.LUpdate.Caption','Recherche de mise à jour');
    AboutBox.UrlUpdate:=Format(BaseUpdateURl,[Version,LangStr]);

    //Accounts
    FAccounts.BtnOk.Caption:=OKBtn;
    FAccounts.BtnCancel.Caption:=CancelBtn;
    FAccounts.LAccName.Caption:=ReadString(LangStr,'FAccounts.LAccName.Caption',FAccounts.LAccName.Caption);
    FAccounts.LHost.Caption:=ReadString(LangStr,'FAccounts.LHost.Caption',FAccounts.LHost.Caption);
    FAccounts.LProtocol.Caption:=ReadString(LangStr,'FAccounts.LProtocol.Caption', FAccounts.LProtocol.Caption);
    FAccounts.CBProtocol.Items[0]:=ReadString(LangStr,'FAccounts.CBProtocol.Items_0',FAccounts.CBProtocol.Items[0]);
    FAccounts.LUserName.Caption:=ReadString(LangStr,'FAccounts.LUserName.Caption',FAccounts.LUserName.Caption);
    FAccounts.LPassword.Caption:=ReadString(LangStr,'FAccounts.LPassword.Caption',FAccounts.LPassword.Caption);
    FAccounts.LEmail.Caption:=ReadString(LangStr,'FAccounts.LEmail.Caption',FAccounts.LEmail.Caption);
    FAccounts.LColor.Caption:=ReadString(LangStr,'FAccounts.LColor.Caption',FAccounts.LColor.Caption);
    //FAccounts.LMailClient.Caption:=ReadString(LangStr,'FAccounts.LMailClient.Caption',FAccounts.LMailClient.Caption);
    FAccounts.LSoundFile.Caption:=ReadString(LangStr,'FAccounts.LSoundFile.Caption',FAccounts.LSoundFile.Caption);
    FAccounts.LSSL.Caption:=ReadString(LangStr,'FAccounts.LSSL.Caption',FAccounts.LSSL.Caption);
    FAccounts.CBSSL.Items[0]:=ReadString(LangStr,'FAccounts.CBSSL.Items_0',FAccounts.CBSSL.Items[0]);
    FAccounts.CBSSL.Items[1]:=ReadString(LangStr,'FAccounts.CBSSL.Items_1',FAccounts.CBSSL.Items[1]);
    FAccounts.CBSSL.Items[2]:=ReadString(LangStr,'FAccounts.CBSSL.Items_2',FAccounts.CBSSL.Items[2]);
    FAccounts.LPort.Caption:=ReadString(LangStr,'FAccounts.LPort.Caption',FAccounts.LPort.Caption);
    FAccounts.CBSecureAuth.Caption:=ReadString(LangStr,'FAccounts.CBSecureAuth.Caption',FAccounts.CBSecureAuth.Caption);
    FAccounts.LInterval.Caption:=ReadString(LangStr,'FAccounts.LInterval.Caption',FAccounts.LInterval.Caption);
    FAccounts.CBShowPass.Caption:=ReadString(LangStr,'FAccounts.CBShowPass.Caption',FAccounts.CBShowPass.Caption);
    FAccounts.LReply.Caption:=ReadString(LangStr,'FAccounts.LReply.Caption',FAccounts.LReply.Caption);
    FAccounts.CBEnabledAcc.Caption:=ReadString(LangStr,'FAccounts.CBEnabledAcc.Caption',FAccounts.CBEnabledAcc.Caption);
    //FAccounts.BtnMailClient.Hint:=ReadString(LangStr,'FAccounts.BtnMailClient.Hint',FAccounts.BtnMailClient.Hint);
    FAccounts.BtnPlaySound.Hint:=ReadString(LangStr,'FAccounts.BtnPlaySound.Hint',FAccounts.BtnPlaySound.Hint);
    FAccounts.BtnSoundFile.Hint:=ReadString(LangStr,'FAccounts.BtnSoundFile.Hint',FAccounts.BtnSoundFile.Hint);

    //Settings
    FSettings.BtnOK.Caption:=OKBtn;
    FSettings.BtnCancel.Caption:=CancelBtn;
    FSettings.Caption:=ReadString(LangStr,'FSettings.Caption',FSettings.Caption);
    FSettings.GBSystem.Caption:=ReadString(LangStr,'FSettings.GBSystem.Caption',FSettings.GBSystem.Caption);
    FSettings.CBStartup.Caption:=ReadString(LangStr,'FSettings.CBStartup.Caption',FSettings.CBStartup.Caption);
    FSettings.CBStartMini.Caption:=ReadString(LangStr,'FSettings.CBStartMini.Caption',FSettings.CBStartMini.Caption);
    FSettings.CBSavSizePos.Caption:=ReadString(LangStr,'FSettings.CBSavSizePos.Caption',FSettings.CBSavSizePos.Caption);
    FSettings.CBMailClientMini.Caption:=ReadString(LangStr,'FSettings.CBMailClientMini.Caption',FSettings.CBMailClientMini.Caption);
    FSettings.CBRestNewMsg.Caption:=ReadString(LangStr,'FSettings.CBRestNewMsg.Caption',FSettings.CBRestNewMsg.Caption);
    FSettings.CBHideInTaskBar.Caption:=ReadString(LangStr,'FSettings.CBHideInTaskBar.Caption',FSettings.CBHideInTaskBar.Caption);
    FSettings.CBSaveLogs.Caption:=ReadString(LangStr,'FSettings.CBSaveLogs.Caption',FSettings.CBSaveLogs.Caption);
    FSettings.CBNoChkNewVer.Caption:=ReadString(LangStr,'FSettings.CBNoChkNewVer.Caption',FSettings.CBNoChkNewVer.Caption);
    FSettings.CBStartupCheck.Caption:=ReadString(LangStr,'FSettings.CBStartupCheck.Caption',FSettings.CBStartupCheck.Caption);
    FSettings.CBSmallBtns.Caption:=ReadString(LangStr,'FSettings.CBSmallBtns.Caption',FSettings.CBSmallBtns.Caption);
    FSettings.CBNotifications.Caption:=ReadString(LangStr,'FSettings.CBNotifications.Caption',FSettings.CBNotifications.Caption);
    FSettings.CBNoCloseAlert.Caption:=ReadString(LangStr,'FSettings.CBNoCloseAlert.Caption',FSettings.CBNoCloseAlert.Caption);
    FSettings.CBNoQuitAlert.Caption:=ReadString(LangStr,'FSettings.CBNoQuitAlert.Caption',FSettings.CBNoQuitAlert.Caption);
    FSettings.LMailClient.Caption:=ReadString(LangStr,'FSettings.LMailClient.Caption',FSettings.LMailClient.Caption);
    FSettings.BtnMailClient.Hint:=ReadString(LangStr,'FSettings.BtnMailClient.Hint',FSettings.BtnMailClient.Hint);
    FSettings.LSoundFile.Caption:=FAccounts.LSoundFile.Caption;
    FSettings.LLangue.Caption:=ReadString(LangStr,'FSettings.LLangue.Caption',FSettings.LLangue.Caption);
    FSettings.BtnPlaySound.Hint:=ReadString(LangStr,'FSettings.BtnPlaySound.Hint',FSettings.BtnPlaySound.Hint);
    FSettings.BtnSoundFile.Hint:=FAccounts.BtnSoundFile.Hint;
    FSettings.CBUrl.Hint:=ReadString(LangStr,'FSettings.CBUrl.Hint',FSettings.CBUrl.Hint);
    FSettings.GMailWeb:=ReadString(LangStr,'FSettings.GMailWeb','Site Web de GMail');
    FSettings.OutlookWeb:=ReadString(LangStr,'FSettings.OutlookWeb','Site Web d''Outlook.com');
    FSettings.Win10Mail:=ReadString(LangStr,'FSettings.Win10Mail','Application Courrier de Windows 10');

    // Choose mail client
    FMailClientChoose.BtnOK.Caption:=OKBtn;
    FMailClientChoose.BtnCancel.Caption:=CancelBtn;
    FMailClientChoose.Caption:=ReadString(LangStr,'FMailClientChoose.Caption',FMailClientChoose.Caption);
    FMailClientChoose.LName.Caption:=ReadString(LangStr,'FMailClientChoose.LName.Caption',FMailClientChoose.LName.Caption);
    FMailClientChoose.LCommand.Caption:=ReadString(LangStr,'FMailClientChoose.LCommand.Caption',FMailClientChoose.LCommand.Caption);
    FMailClientChoose.CBUrl.Hint:=FSettings.CBUrl.Hint;
    FMailClientChoose.BtnMailClient.Hint:=FSettings.BtnMailClient.Hint;

    // Impex
    FImpex.BtnOK.Caption:=OKBtn;
    FImpex.BtnCancel.Caption:=CancelBtn;
    FImpex.LAccTyp.Caption:=ReadString(LangStr,'FImpex.LAccTyp.Caption',FImpex.LAccTyp.Caption);
    FImpex.Caption:=ReadString(LangStr,'FImpex.Caption',FImpex.Caption);
    FImpex.LFilename.Caption:=ReadString(LangStr,'FImpex.LFilename.Caption',FImpex.LFilename.Caption);
    Fimpex.MailattAccName:=ReadString(LangStr,'Fimpex.MailattAccName', 'Comptes MailAttente');
    Fimpex.OutlAccName:=ReadString(LangStr,'Fimpex.OutlAccName', 'Comptes Outlook 2007-2010');
    Fimpex.SGImpex.Columns[0].Title.Caption:=ReadString(LangStr,'Fimpex.SGImpex.Columns_0.Title.Caption',
           Fimpex.SGImpex.Columns[0].Title.Caption) ;
    Fimpex.SGImpex.Columns[1].Title.Caption:=ReadString(LangStr,'Fimpex.SGImpex.Columns_1.Title.Caption',
           Fimpex.SGImpex.Columns[0].Title.Caption) ;
    Fimpex.SGImpex.Cells[0,1]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_1','Nom du compte');
    Fimpex.SGImpex.Cells[0,2]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_2','Serveur de courrier');
    Fimpex.SGImpex.Cells[0,3]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_3','Port');
    Fimpex.SGImpex.Cells[0,4]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_4','Protocole');
    Fimpex.SGImpex.Cells[0,5]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_5','Identifiant courriel');
    Fimpex.SGImpex.Cells[0,6]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_6','Mot de passe');
    Fimpex.SGImpex.Cells[0,7]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_7','Adresse courriel');
    Fimpex.SGImpex.Cells[0,8]:=ReadString(LangStr,'Fimpex.SGImpex.Cells_0_8','Adresse de réponse');
    Fimpex.spassNotAvail:=ReadString(LangStr,'Fimpex.spassNotAvail','Mot de passe non disponible');
    Fimpex.ODImpex.Title:=ReadString(LangStr,'Fimpex.ODImpex.Title',Fimpex.ODImpex.Title);
    FImpex.xmlFilter:=ReadString(LangStr,'FImpex.xmlFilter','Fichiers XML|*.xml|Tous les fichiers|*.*');
    FImpex.jsFilter:=ReadString(LangStr,'FImpex.jsFilter','Fichiers Javascript|*.js|Tous les fichiers|*.*');
    FImpex.sBtnAccFileHint:=ReadString(LangStr,'FImpex.sBtnAccFileHint',FImpex.BtnAccFile.Hint);

    // Tray
    MnuRestore.Caption:=ReadString(LangStr,'MnuRestore.Caption',MnuRestore.Caption);
    MnuMaximize.Caption:=ReadString(LangStr,'MnuMaximize.Caption',MnuMaximize.Caption);
    MnuIconize.Caption:=ReadString(LangStr,'MnuMinimize.Caption',MnuIconize.Caption);
    MnuGetAllMail.Caption:= BtnGetAllMail.Hint;
    MnuQuit.Caption:= ReadString(LangStr,'MnuQuit.Caption',MnuQuit.Caption);
    MnuAbout.Caption:=BtnAbout.Hint;

        // HTTP Error messages
    HttpErrMsgNames[0] := ReadString(LangStr,'SErrInvalidProtocol','Protocole "%s" invalide');
    HttpErrMsgNames[1] := ReadString(LangStr,'SErrReadingSocket','Erreur de lecture des données à partir du socket');
    HttpErrMsgNames[2] := ReadString(LangStr,'SErrInvalidProtocolVersion','Version de protocole invalide en réponse: %s');
    HttpErrMsgNames[3] := ReadString(LangStr,'SErrInvalidStatusCode','Code de statut de réponse invalide: %s');
    HttpErrMsgNames[4] := ReadString(LangStr,'SErrUnexpectedResponse','Code de statut de réponse non prévu: %s');
    HttpErrMsgNames[5] := ReadString(LangStr,'SErrChunkTooBig','Bloc trop grand');
    HttpErrMsgNames[6] := ReadString(LangStr,'SErrChunkLineEndMissing','Fin de ligne du bloc manquante');
    HttpErrMsgNames[7] := ReadString(LangStr,'SErrMaxRedirectsReached','Nombre maximum de redirections atteint: %s');
    // Socket error messages
    HttpErrMsgNames[8] := ReadString(LangStr,'strHostNotFound','Résolution du nom d''hôte pour "%s" impossible.');
    HttpErrMsgNames[9] := ReadString(LangStr,'strSocketCreationFailed','Echec de la création du socket: %s');
    HttpErrMsgNames[10] := ReadString(LangStr,'strSocketBindFailed','Echec de liaison du socket: %s');
    HttpErrMsgNames[11] := ReadString(LangStr,'strSocketListenFailed','Echec de l''écoute sur le port n° %s, erreur %s');
    HttpErrMsgNames[12]:=ReadString(LangStr,'strSocketConnectFailed','Echec de la connexion à %s');
    HttpErrMsgNames[13]:=ReadString(LangStr,'strSocketAcceptFailed','Connexion refusée d''un client sur le socket: %s, erreur %s');
    HttpErrMsgNames[14]:=ReadString(LangStr,'strSocketAcceptWouldBlock','La connexion pourrait bloquer le socket: %s');
    HttpErrMsgNames[15]:=ReadString(LangStr,'strSocketIOTimeOut','Impossible de fixer le timeout E/S à %s');
    HttpErrMsgNames[16]:=ReadString(LangStr,'strErrNoStream','Flux du socket non assigné');

  end;
end;

end.

