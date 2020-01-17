{******************************************************************************}
{ MailInBox main unit                                                          }
{ bb - sdtp - january 2020                                                     }
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
  lazbbinifiles, lazbbosversion, LazUTF8, settings1, lazbbautostart, lazbbabout,
  Impex1, mailclients1, uxtheme, Types, IdComponent, fptimer, lazbbalert,
  IdMessageCollection, RichMemo, UniqueInstance, csvdocument, log1, registry;

type
  TSaveMode = (None, Setting, All);
  TBtnSize=(Small, Large);


  { TFMailsInBox }
  TFMailsInBox = class(TForm)
    BtnAbout: TSpeedButton;
    BtnAddAcc: TSpeedButton;
    BtnClose: TSpeedButton;
    BtnDelete: TSpeedButton;
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
    ImgAccounts: TImageList;
    LNow: TLabel;
    LStatus: TLabel;
    LVAccounts: TListView;
    MenuItem1: TMenuItem;
    MenuItem3: TMenuItem;
    MnuAnswerMsg: TMenuItem;
    MnuDeleteMsg: TMenuItem;
    MnuGetAllMail: TMenuItem;
    MnuIconize: TMenuItem;
    MenuItem2: TMenuItem;
    MnuAbout: TMenuItem;
    MnuMaximize: TMenuItem;
    MnuQuit: TMenuItem;
    MnuRestore: TMenuItem;
    MnuInfos: TMenuItem;
    MnuMoveDown: TMenuItem;
    MnuMoveUp: TMenuItem;
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
    RMInfos: TRichMemo;
    SplitterV: TSplitter;
    SplitterH: TSplitter;
    SGMails: TStringGrid;
    GetMailTimer: TTimer;
    TrayTimer: TTimer;
    TrayMail: TTrayIcon;
    UniqueInstance1: TUniqueInstance;
    procedure BtnAboutClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);
    procedure BtnGetAccMailClick(Sender: TObject);
    procedure BtnGetAllMailClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure BtnLaunchClientClick(Sender: TObject);
    procedure BtnLogClick(Sender: TObject);
    procedure BtnQuitClick(Sender: TObject);
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
    procedure MnuAboutClick(Sender: TObject);
    procedure MnuAccountPopup(Sender: TObject);
    procedure MnuAnswerMsgClick(Sender: TObject);
    procedure MnuDeleteMsgClick(Sender: TObject);
    procedure MnuInfosClick(Sender: TObject);
    procedure MnuMailsPopup(Sender: TObject);
    procedure MnuMaximizeClick(Sender: TObject);
    procedure MnuIconizeClick(Sender: TObject);
    procedure MnuMoveDownClick(Sender: TObject);
    procedure MnuMoveUpClick(Sender: TObject);
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
    procedure TrayTimerTimer(Sender: TObject);
    procedure OnChkMailTimer(Sender: TObject);
  private
    Initialized: boolean;
    OS, OSTarget, CRLF: string;
    CompileDateTime: TDateTime;
    UserPath, MIBAppDataPath: string;
    ProgName: string;
    LangStr: string;
    LangFile: TBbIniFile;
    LangNums: TStringList;
    LangFound: boolean;
    SettingsChanged: boolean;
    AccountsChanged:Boolean;
    ConfigFileName: string;
    ChkMailTimerTick: integer;
    CanClose: boolean;
    BaseUpdateUrl, ChkVerURL, version: string;
    AccountCaption, EmailCaption, LastCheckCaption, NextCheckCaption : string;
    sNoLongerChkUpdates, sLastUpdateSearch, sUpdateAvailable, sUpdateAlertBox: string;
    OKBtn, YesBtn, NoBtn, CancelBtn: string;
    BtnLogHint, BtnGetAccMailHint, BtnDeleteHint, BtnEditAccHint: string;
    BmpArray: array of TBitmap;
    sAccountImported, sAccountsImported: string;
    CheckingMail: boolean;
    SGHasFocus: boolean;
    DefCursor: TCursor;
    sMsgFound, sMsgsFound: string;
    ChkMailTimer: TFPTimer;
    TrayTimerTick: integer;
    TrayTimerBmp: TBitmap;
    LStatusCaption: String;
    MainLog: String;
    SessionLog: TStringList;
    CurAccPend: integer;
    sCheckingAccMail, sCheckingAllMail: string;
    sConnectToServer, sDisconnectServer: string;
    sConnectedToServer, DisconnectedServer: string;
    sMsgDeleted, sMsgNotDeleted: string;
    ConnectErrorMsg, HeaderErrorMsg: string;
    sAccountChanged, sAccountAdded: string;
    sAccountDisabled, sAccountEnabled, AccountStatus: string;
    LogFileName: string;
    sLoadingAccounts: string;
    Iconized: Boolean;
    PrevTop, PrevLeft: integer;
    sTrayHintNoMsg, sTrayHintMsg, sTrayHintMsgs: string;
    sTrayHintNewMsg, sTrayHintNewMsgs: string;
    sTrayBallHintMsg, sTrayBallHintMsgs: string;
    sNoCloseAlert, sNoQuitAlert, sNoShowCloseAlert, sNoShowQuitAlert: string;
    sColumnswidth: string;
    sOpenProgram, sRestart: string;
    sRetConfBack, sCreNewConf, sLoadConf: string;
    sCannotQuit, sClosingProg: string;
    sSettingsChange: string;
    StatusFmtSets: TFormatSettings;
    sMnuDelMsg, sMnuAnswerMsg: string;
    sAlertDelMmsg: string;
    AccountPictures, TrayPicture, MailPictures: TPicture;
    procedure Initialize;
    procedure LoadCfgFile(filename: string);
    procedure SettingsOnChange(Sender: TObject);
    procedure SettingsOnStateChange(Sender: TObject);
    procedure AccountsOnChange(Sender: TObject);
    function SaveConfig(Typ: TSaveMode): boolean;
    procedure PopulateAccountsList(notify: boolean);
    procedure ModLangue;
    procedure SetSmallBtns(small: boolean);
    function GetPendingMail(index: integer): integer;
    procedure PopulateMailsList(index: integer);
    procedure EnableControls(Enable: boolean);
    procedure UpdateInfos;
    procedure DrawTheIcon(Bmp: TBitmap; NewCount: integer; CircleColor: TColor);
    function MailChecking(status: boolean): boolean;
    procedure GetMailInfos(CurName: String; var Mail: TMail; IdMsg: TIdMessage; siz: Integer);
    function LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
    procedure GetAccMail(ndx: integer);
    function HideOnTaskbar: boolean;
    procedure OnAppMinimize(Sender: TObject);
    procedure OnQueryendSession(var Cancel: Boolean);
    procedure GetMnuImage(il: TImageList; Resname: string; twin: boolean=true);
  public
    OsInfo: TOSInfo;           //used by Settings1
    UserAppsDataPath: string;  //used by Impex1
  end;

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
  // Initialize timers stuff
  ChkMailTimer:= TFPTimer.Create(self);
  ChkMailTimer.Interval:= 100;
  ChkMailTimer.UseTimerThread:= true;
  ChkMailTimer.Enabled:= true;
  ChkMailTimer.OnTimer:= @OnChkMailTimer;
  ChkMailTimerTick:= 0;
  // Intercept minimize system command
  Application.OnMinimize:=@OnAppMinimize;
  Application.OnQueryEndSession:= @OnQueryendSession;
  TrayTimerTick:=0;
  TrayTimerBmp:= TBitmap.Create;
  // Flag needed to execute once some processes in Form activation
  Initialized := False;
  MainLog:= '';
  SessionLog:= TStringList.Create;
  CompileDateTime:= StringToTimeDate({$I %DATE%}+' '+{$I %TIME%}, 'yyyy/mm/dd hh:nn:ss');
  OS := 'Unk';
  UserPath := GetUserDir;
  UserAppsDataPath := UserPath;
  {$IFDEF Linux}
    OS := 'Linux';
    CRLF := #10;
    LangStr := GetEnvironmentVariable('LANG');
    x := pos('.', LangStr);
    LangStr := Copy(LangStr, 0, 2);
    wxbitsrun := 0;
    OSTarget:= '';
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
  ProgName := 'MailsInBox';
  version := GetVersionInfo.ProductVersion;
  // Chargement des chaînes de langue...
  LangFile := TBbIniFile.Create(ExtractFilePath(Application.ExeName) + LowerCase(ProgName)+'.lng');
  // Cannot call Modlang as components are not yet created, use default language
  sOpenProgram:=LangFile.ReadString(LangStr,'OpenProgram','Ouverture de Courrier en attente');
  LogAddLine(-1, now, sOpenProgram+' - Version '+Version+ ' (' + OS + OSTarget + ')');
  LogAddLine(-1, now, OsInfo.VerDetail);

  LangNums := TStringList.Create;
  MIBAppDataPath := UserAppsDataPath + PathDelim + ProgName + PathDelim;
  if not DirectoryExists(MIBAppDataPath) then CreateDir(MIBAppDataPath);
  LogFileName:= MIBAppDataPath+ProgName+'.log';
  // Mail checking process flag
  CheckingMail:= false;
  iconized:= false;
  // DateTime settings for status bar
  StatusFmtSets:= DefaultFormatSettings ;
  StatusFmtSets.ShortDateFormat:= DefaultFormatSettings.LongDateFormat;
end;

// Add a line to log file. quote char is '|' to preserve real double quotes in log
// when reading with csv reader

function TFMailsInBox.LogAddLine(acc: integer; dat: TDateTime; evnt: string): integer;
begin
 result:= SessionLog.Add('|'+Inttostr(acc)+'|,'+'|'+TimeDateToString(dat)+'|,'+
                      '|'+evnt+'|');
end;

// Form activation only needed once

procedure TFMailsInBox.FormActivate(Sender: TObject);
begin
  Initialize;
end;

// populate menu imagelist from resource name(resname)
// Twin= true : image for enabled and disabled item

procedure TFMailsInBox.GetMnuImage(il: TImageList; Resname: string; twin: boolean=true);
var
  Pict: TPicture;
  bmp: TBitmap;
begin
  Pict:= TPicture.Create;
  Bmp:=Tbitmap.create;
  Pict.LoadFromResourceName(HInstance, Resname);
  CropBitmap(Pict.Bitmap, Bmp, true);
  il.AddMasked(Bmp, $FF00FF);
  // We need disabled image
  if twin then
  begin
    CropBitmap(Pict.Bitmap, Bmp, false);
    il.AddMasked(Bmp, $FF00FF);
  end;
  if Assigned(Bmp) then Bmp.Free;
  if Assigned(Pict) then Pict.Free;
end;

// Initializing stuff

procedure TFMailsInBox.Initialize;
var
  i: integer;
  defmailcli: string;
  IniFile: TBbIniFile;
  tmplog: TstringList;
  CurAcc: TAccount;
begin
  if initialized then exit;
  // For popup menu, retrieve bitmap from resource
  // Get tray menu images from ressource
  ILMnuTray.Clear;
  GetMnuImage(ILMnuTray, 'RESTORE16');
  GetMnuImage(ILMnuTray, 'MAXIMIZE16');
  GetMnuImage(ILMnuTray, 'ICONIZE16');
  GetMnuImage(ILMnuTray, 'GETALLMAIL16');
  GetMnuImage(ILMnuTray, 'ABOUT16');
  GetMnuImage(ILMnuTray, 'QUIT16');
  ILMnuMails.Clear;
  GetMnuImage(ILMnuMails, 'MAILINFOS16');
  GetMnuImage(ILMnuMails, 'MAILANSWER16');
  GetMnuImage(ILMnuMails, 'MAILDELETE16');
  ILMnuAccounts.Clear;
  GetMnuImage(ILMnuAccounts, 'ARROWUP16');
  GetMnuImage(ILMnuAccounts, 'ARROWDN16');
  // Mail list images
  AccountPictures:= TPicture.Create;
  AccountPictures.LoadFromResourceName(HInstance, 'ACCOUNT216');
  MailPictures:= TPicture.Create;
  MailPictures.LoadFromResourceName(HInstance, 'MAILSTATES');
  // Tray icon
  TrayPicture:= Tpicture.Create;
  TrayPicture.LoadFromResourceName(HInstance, 'MAIL16');
  // Now, main settings
  FSettings.Settings.AppName:= LowerCase(ProgName);
  FAccounts.Accounts.AppName := LowerCase(ProgName);
  ConfigFileName := MIBAppDataPath + ProgName + '.xml';
  ModLangue;
  if not FileExists(ConfigFileName) then
  begin
    if FileExists(MIBAppDataPath + ProgName + '.bk0') then
    begin
      LogAddLine(-1, now, sRetConfBack);
      RenameFile(MIBAppDataPath + ProgName + '.bk0', ConfigFileName);
      for i := 1 to 5 do
        if FileExists(MIBAppDataPath + ProgName + '.bk' + IntToStr(i))
        // Renomme les précédentes si elles existent
        then
          RenameFile(MIBAppDataPath + ProgName + '.bk' + IntToStr(i),
            MIBAppDataPath + ProgName + '.bk' + IntToStr(i - 1));
    end else
    begin
      SaveConfig(All);
      LogAddLine(-1, now, sCreNewConf);
    end;
  end;
  LoadCfgFile(ConfigFileName);
  // Check inifile with URLs, if not present, then use default
  IniFile:= TBbInifile.Create('mailsinbox.ini');
  BaseUpdateUrl:= IniFile.ReadString('urls', 'BaseUpdateUrl',
    'https://www.sdtp.com/versions/version.php?program=mailsinbox&version=%s&language=%s');
  ChkVerURL := IniFile.ReadString('urls', 'ChkVerURL',
    'https://www.sdtp.com/versions/versions.csv');
  if Assigned(IniFile) then IniFile.free;
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
  AlertBox.Caption:= Caption;
  //AlertBox.Image1.Picture.LoadFromResourceName(HInstance, 'ABOUTIMG');
  // Load last log file
  tmplog:= TStringList.Create;
  if FileExists(LogFileName) then
  begin
    tmpLog.LoadFromFile(LogFileName);
    MainLog:= tmpLog.text;
  end;
  BtnLaunchClient.Enabled:= not (FSettings.Settings.MailClient='');
  // Accounts are sorted on their index
  FAccounts.Accounts.SortType:= cdcindex;
  FAccounts.Accounts.DoSort;
  PopulateAccountsList(false);
  if  FAccounts.Accounts.count > 0 then LVAccounts.ItemIndex:= 0
  else LVAccounts.ItemIndex:=  -1;
  FSettings.Settings.OnChange := @SettingsOnChange;
  FSettings.Settings.OnStateChange := @SettingsOnStateChange;
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
  //Save large buttons glyphs in an array
  SetLength(BmpArray, PnlToolbar.ControlCount);
  for i:=0 to PnlToolbar.ControlCount-1 do
  begin
    BmpArray[i]:= TBitmap.create;
    BmpArray[i].Assign(TSpeedButton(PnlToolbar.Controls[i]).Glyph);
  end;
  // set small buttons only if asked
  if FSettings.Settings.SmallBtns then SetSmallBtns(FSettings.Settings.SmallBtns);
  Constraints.MinWidth:= BtnQuit.left+BtnQuit.width+10;
  if width < Constraints.MinWidth then width := Constraints.MinWidth;
  // Get default mail client and stores its name
  defmailcli:= FSettings.GetDefaultMailCllient;
  // Launch config dailog to set mail client and other stuff
  if length(FSettings.Settings.MailClient)=0 then
  begin
    FSettings.Settings.MailClient:= defmailcli;
    BtnSettingsClick(self);
  end;
  LogAddLine(-1, now, 'Mail Client: '+FSettings.Settings.MailClientName);
  // TStringgrid MousetoCell give cell nearest the mouse click if
  // this property is true. To get -1 when mouse is outside a cell then
  SGMails.AllowOutboundEvents:=false;
  TrayMail.Hint:= sTrayHintNoMsg;
  if FSettings.Settings.StartupCheck then BtnGetAllMailClick(self);
  GetMailTimer.enabled:= true;
  Initialized:= true;
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
  SettingsChanged:= FSettings.Settings.SavSizePos;
end;

// change size of buttons in toolbar

procedure TFMailsInBox. SetSmallBtns(small:boolean);
var
  Pict:TPicture;
  i: integer;
  ImgName:string;
begin
  if Small then PnlToolbar.height:= 34
  else PnlToolbar.height:= 42;
  Pict:= TPicture.Create;
  for i:= 0 to PnlToolbar.ControlCount-1 do
  begin
    ImgName:= uppercase(TSpeedButton(PnlToolbar.Controls[i]).name);
    ImgName:=copy(ImgName,4, length(ImgName)-3);
    if Small then
      begin
        if PnlToolbar.Controls[i].ClassType=TSpeedButton then
      try
        TSpeedButton(PnlToolbar.Controls[i]).Height:= 24;
        TSpeedButton(PnlToolbar.Controls[i]).Width:=24;
        Pict.LoadFromResourceName(HInstance, ImgName+'16');
        TSpeedButton(PnlToolbar.Controls[i]).Left:= 10+i*(24+8);
        TSpeedButton(PnlToolbar.Controls[i]).Glyph.Assign(Pict.Bitmap);
      except
      end;
    end else
      begin
        TSpeedButton(PnlToolbar.Controls[i]).Height:= 32;
        TSpeedButton(PnlToolbar.Controls[i]).Width:=32;
        TSpeedButton(PnlToolbar.Controls[i]).Left:= 10+i*(32+8);
        TSpeedButton(PnlToolbar.Controls[i]).Glyph.Assign(BmpArray[i]);
      end;
  end;
  Pict.Free;
end;

// Load configuration and database from file

procedure TFMailsInBox.LoadCfgFile(filename: string);
var
  winstate: TWindowState;
  i: integer;
begin
  with FSettings do
  begin
    LogAddLine(-1, now, sLoadConf);
    Settings.LoadXMLFile(filename);
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
    // Get columns width to use at application close
    sColumnswidth:='';
    For i:= 0 to 4 do  sColumnswidth:= sColumnswidth+IntToHex(self.SGMails.Columns [i].Width, 4);
    if settings.StartMini then Application.Minimize;
    // Détermination de la langue (si pas dans settings, langue par défaut)
    if Settings.LangStr = '' then Settings.LangStr := LangStr;
    LangFile.ReadSections(LangNums);
    if LangNums.Count > 1 then
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
    // Si la langue n'est pas traduite, alors on passe en Anglais
    if not LangFound then
    begin
      Settings.LangStr := 'en';
    end;
  end;
  Modlangue;
  Application.Title:=Caption;
  LogAddLine(-1, now, sLoadingAccounts);
  FAccounts.Accounts.Reset;
  FAccounts.Accounts.LoadXMLfile(filename);
  SettingsChanged := false;
end;

procedure TFMailsInBox.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  s: string;
  i: integer;
  curcolsw: string;
begin
  if not visible then AlertBox.Position:= poDesktopCenter
  else AlertBox.Position:= poMainFormCenter;
  // Alert box to explain close and quit operation
  //if CanClose then
  //else AlertBox.MAlert.Text:= sNoCloseAlert;
  AlertBox.DlgType:= mtInformation;

    {  end else
  begin
  if not FSettings.Settings.NoCloseAlert then
  begin
    AlertBox.CBNoShowAlert.Checked:= FSettings.Settings.NoCloseAlert;
    if AlertBox.ShowModal=mrOK then
    begin
      if AlertBox.CBNoShowAlert.Checked then
      begin
        FSettings.Settings.NoCloseAlert:=true;
        LogAddLine(-1, now, sNoShowCloseAlert);
      end;
    end;

  end;
     CloseAction:= caNone;
     exit;
   end;    }
  if CanClose then
  begin
    if not FSettings.Settings.NoQuitAlert then
    begin
      AlertBox.MAlert.Text:= sNoQuitAlert;
      AlertBox.CBNoShowAlert.Checked:= FSettings.Settings.NoQuitAlert;
      if AlertBox.ShowModal=mrOK then
      begin
        if AlertBox.CBNoShowAlert.Checked then
        begin
          FSettings.Settings.NoQuitAlert:=true;
          LogAddLine(-1, now, sNoShowQuitAlert);
        end;
      end else
      begin
        CanClose:= false;
        CloseAction:= caNone;
        MnuIconizeClick(sender);
        exit;
      end;
    end;
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
    if not FSettings.Settings.NoCloseAlert then
    begin
      AlertBox.MAlert.Text:= sNoCloseAlert;
      AlertBox.CBNoShowAlert.Checked:= FSettings.Settings.NoCloseAlert;
      if AlertBox.ShowModal=mrOK then
      begin
        if AlertBox.CBNoShowAlert.Checked then
        begin
          FSettings.Settings.NoCloseAlert:=true;
          LogAddLine(-1, now, sNoShowCloseAlert);
        end;
      end;
    end;
    MnuIconizeClick(sender);
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
    if FileExists (ConfigFileName) then
    begin
      if (Typ = All) then
      begin
        // On sauvegarde les versions précédentes parce que la base de données a changé
        FilNamWoExt:= TrimFileExt(ConfigFileName);
        if FileExists (FilNamWoExt+'.bk5')                   // Efface la plus ancienne
        then  DeleteFile(FilNamWoExt+'.bk5');                // si elle existe
        For i:= 4 downto 0
        do if FileExists (FilNamWoExt+'.bk'+IntToStr(i))     // Renomme les précédentes si elles existent
           then  RenameFile(FilNamWoExt+'.bk'+IntToStr(i), FilNamWoExt+'.bk'+IntToStr(i+1));
        RenameFile(ConfigFileName, FilNamWoExt+'.bk0');
        FAccounts.Accounts.SaveToXMLfile(ConfigFileName);
      end;
      // la base n'a pas changé, on ne fait pas de backup
      FSettings.settings.SaveToXMLfile(ConfigFileName);
    end else
    begin
      FAccounts.Accounts.SaveToXMLfile(ConfigFileName);
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
  if FAccounts.Accounts.Count = 0 then exit;
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

procedure TFMailsInBox.BtnEditAccClick(Sender: TObject);
var
  Account: TAccount;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx <0 then exit;
  with FAccounts do
  begin
    if (TSpeedButton(Sender).Name='BtnAddAcc') then
    begin
      Account:= Default(TAccount);
      Account.UID:=0;
      Caption:= BtnAddAcc.Hint;
    end;
    if (TSpeedButton(Sender).Name='BtnEditAcc') and  (ndx>=0) then
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
  for i:=0 to length(BmpArray)-1 do if Assigned(BmpArray[i]) then BmpArray[i].free;
  if Assigned(ChkMailTimer) then ChkMailTimer.Destroy;
  if Assigned(SessionLog) then SessionLog.free;
  if Assigned(AccountPictures) then AccountPictures.Free;
  if Assigned(TrayPicture) then TrayPicture.Free;
  if Assigned(MailPictures) then MailPictures.Free;
end;

// Timer firing periodic mail checking

procedure TFMailsInBox.GetMailTimerTimer(Sender: TObject);
var
  i: integer;
  min: TDateTime;
  CurAcc: TAccount;
   ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx <0 then exit;
  for i:=0 to FAccounts.Accounts.count-1 do
  begin
    // current account is enabled and interval defined
    CurAcc:= FAccounts.Accounts.GetItem(i);
    if CurAcc.Enabled and (CurAcc.Interval>0) and (now>CurAcc.NextFire) then
    begin
      min:= EncodeTime(0,CurAcc.interval,0,0);
      GetAccMail(i);
      FAccounts.Accounts.ModifyField(i, 'NEXTFIRE', now+min);
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
  LStatus.Caption:= Format(DisconnectedServer, [Curacc.Name, CurAcc.Server]);
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
    LVAccounts.PopupMenu:= MnuAccount;
    AccName:= FAccounts.Accounts.GetItem(ndx).Name;
    BtnAccountLog.Hint:= Format(BtnLogHint, [AccName]);
    BtnGetAccMail.Hint:= Format(BtnGetAccMailHint, [AccName]);
    BtnDelete.Hint:= Format(BtnDeleteHint, [AccName]);
    BtnEditAcc.Hint:= Format(BtnEditAccHint, [AccName]);
    UpdateInfos;
  end else LVAccounts.PopupMenu:= nil;
  PopulateMailsList(ndx);
end;



procedure TFMailsInBox.MnuAboutClick(Sender: TObject);
begin

end;

procedure TFMailsInBox.UpdateInfos;
var
  ndx: integer;
  msgs: integer;
  msgsfnd: string;
  CurAcc: TAccount;
  slastfire: string;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx >= 0 then
  begin
    CurAcc:= FAccounts.Accounts.GetItem(ndx);
    msgs:= CurAcc.Mails.Count;
    slastfire:= TimeDateToString(CurAcc.LastFire);
    RMInfos.Clear;
    if CurAcc.Enabled then RMInfos.Lines.Add(Format(sAccountEnabled, [CurAcc.Name]))
    else RMInfos.Lines.Add(Format(sAccountDisabled, [CurAcc.Name]));
    RMInfos.Lines.Add(Format(EmailCaption, [CurAcc.Email]));
    RMInfos.Lines.Add(Format(LastCheckCaption, [slastfire]));
    if CurAcc.Enabled then
    begin
      if CurAcc.error then RMInfos.Lines.Add(CurAcc.ErrorStr);

      if msgs>1 then msgsfnd:= Format(sMsgsFound, [msgs])
      else msgsfnd:= Format(sMsgFound, [msgs]);
      RMInfos.Lines.Add(msgsfnd);
      LStatus.Caption:= Format(LStatusCaption, [msgsfnd,
                   FAccounts.Accounts.GetItem(ndx).Name, slastfire]);
      RMInfos.Lines.add(Format(NextCheckCaption, [TimeDateToString(FAccounts.Accounts.GetItem(ndx).NextFire)]));
    end else
    begin
      LStatus.Caption:= Format(sAccountDisabled, [CurAcc.Name]);
    end;
   end;
end;

procedure TFMailsInBox.PopulateMailsList(index: integer);
var
  i: integer;
  siz: integer;
  CurAcc: TAccount;
  s: string;
begin
  //SGMails.Clear;
  SGMails.RowCount:=1;
  if (index<0) or (FAccounts.Accounts.GetItem(index).Mails.Count=0) then exit;
  CurAcc:= FAccounts.Accounts.GetItem(index);
  SGMails.RowCount:= CurAcc.Mails.Count+1;
  for i:= 0 to SGMails.RowCount-2 do
  begin
    SGMails.Cells[0,i+1]:= CurAcc.Mails.GetItem(i).MessageFrom;
    SGMails.Cells[1,i+1]:= CurAcc.Mails.GetItem(i).AccountName;
    //SGMails.Cells[1,i+1]:= CurAcc.Mails.GetItem(i).MessageFrom;
    SGMails.Cells[2,i+1]:= CurAcc.Mails.GetItem(i).MessageSubject;
    SGMails.Cells[3,i+1]:= TimeDateToString(CurAcc.Mails.GetItem(i).MessageDate);
    // Change unit with size value
    siz:= CurAcc.Mails.GetItem(i).MessageSize;
    if siz < 20480 then s:= InttoStr(siz)+' bytes'
    else if siz < 100480 then s:= Format('%.1n KB', [siz/1048])
         else if siz < 1048576 then s:= Format('%u KB', [siz div 1048])
              else s:= Format('%.1n MB', [siz/1048576]);
    SGMails.Cells[4,i+1]:= s;
  end;

end;

procedure TFMailsInBox.MnuAccountPopup(Sender: TObject);
begin
  MnuMoveUp.Enabled:= not (LVAccounts.ItemIndex=0);
  if MnuMoveUp.Enabled then MnuMoveUp.ImageIndex:=0 else MnuMoveUp.ImageIndex:=1 ;
  MnuMoveDown.Enabled:= not (LVAccounts.ItemIndex=LVAccounts.Items.count-1);
  if MnuMoveDown.Enabled then MnuMoveDown.ImageIndex:=2 else MnuMoveDown.ImageIndex:=3;
end;

procedure TFMailsInBox.MnuAnswerMsgClick(Sender: TObject);
begin

end;

procedure TFMailsInBox.MnuDeleteMsgClick(Sender: TObject);
var
  andx: integer;
  mndx: integer;
  //mail:TMail;
  uid2del: integer;
  CurAcc: TAccount;
  Subj: string;
begin
  andx:= LVAccounts.ItemIndex;
  if andx<0 then exit;
  mndx:= SGMails.row-1;
  if andx <0 then exit;
  CurAcc:= FAccounts.Accounts.GetItem(andx);
  Subj:= Copy(CurAcc.Mails.GetItem(mndx).MessageSubject,1, 15)+'...';
  // Alarm before deleting
  if MsgDlg(Caption, Format(sAlertDelMmsg, [Subj]), mtWarning,
             [mbYes,mbNo], [YesBtn,NoBtn])= mrYes then
  begin
     // Add our new UID only if not already marked
    if CurAcc.Mails.GetItem(mndx).MessageToDelete then exit;
    TAccount(FAccounts.Accounts.Items[andx]^).Mails.ModifyField(mndx, 'MessageToDelete', true);
    // Add UID to array
    uid2del:= length(CurAcc.UIDLToDel);
    Setlength(TAccount(FAccounts.Accounts.Items[andx]^).UIDLToDel, uid2del+1);
    TAccount(FAccounts.Accounts.Items[andx]^).UIDLToDel[uid2del]:= CurAcc.Mails.GetItem(mndx).MessageUIDL;
    SGMails.Invalidate;
  end;
end;

procedure TFMailsInBox.MnuInfosClick(Sender: TObject);
var
  andx: integer;
  mndx: integer;
  mail:TMail;
  s: string;

begin
  andx:= LVAccounts.ItemIndex;
  if andx <0 then exit;
  mndx:= SGMails.row-1;
  mail:= FAccounts.Accounts.GetItem(andx).Mails.GetItem(mndx);
  s:= 'De : ';
  if length(Mail.MessageFrom)<0 then s:=s+Mail.MessageFrom+' ';
  s:=s+'('+mail.FromAddress+')'+#10;
  s:=s+'A : '+ FAccounts.Accounts.GetItem(andx).Name+'('+Mail.ToAddress+')'+#10;
  s:=s+'Sujet : '+Mail.MessageSubject+#10;
  s:=s+'Date : '+TimeDateToString(Mail.MessageDate)+#10;
  s:=s+'Taille : '+IntToStr(Mail.MessageSize)+#10;
  s:=s+'UID : '+Mail.MessageUIDL;
  ShowMessage(s);
end;

procedure TFMailsInBox.MnuMailsPopup(Sender: TObject);
var
  CurAcc: TAccount;
  Subj: string;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then exit;
  CurAcc:= FAccounts.Accounts.GetItem(ndx);
  Subj:= Copy(CurAcc.Mails.GetItem(SGMails.row-1).MessageSubject,1, 15)+'...';
  MnuDeleteMsg.Caption:= Format(sMnuDelMsg, [Subj]);
  MnuAnswerMsg.Caption:= Format(sMnuAnswerMsg, [Subj]);
  if MnuInfos.Enabled then MnuInfos.ImageIndex:=0 else MnuInfos.ImageIndex:=1 ;
  if MnuDeleteMsg.Enabled then MnuDeleteMsg.ImageIndex:=2 else MnuMaximize.ImageIndex:=3;
  if MnuAnswerMsg.Enabled then MnuAnswerMsg.ImageIndex:=4 else MnuAnswerMsg.ImageIndex:=5 ;
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

procedure TFMailsInBox.MnuMoveDownClick(Sender: TObject);
var
  oldndx: integer;
begin
  oldndx:= LVAccounts.ItemIndex;
  if oldndx<0 then exit;
  if oldndx<LVAccounts.Items.count-1 then
  begin
    FAccounts.Accounts.sorttype:= cdcNone;
    FAccounts.Accounts.ModifyField(oldndx, 'index', oldndx+1);
    FAccounts.Accounts.ModifyField(oldndx+1, 'index', oldndx);
    FAccounts.Accounts.sorttype:= cdcIndex;
    PopulateAccountsList(false);
    LVAccounts.ItemIndex:= oldndx+1;
  end;
end;

procedure TFMailsInBox.MnuMoveUpClick(Sender: TObject);
var
  oldndx: integer;
begin
  oldndx:= LVAccounts.ItemIndex;
  if oldndx>0 then
  begin
    FAccounts.Accounts.sorttype:= cdcNone;
    FAccounts.Accounts.ModifyField(oldndx, 'index', oldndx-1);
    FAccounts.Accounts.ModifyField(oldndx-1, 'index', oldndx);
    FAccounts.Accounts.sorttype:= cdcIndex;
    PopulateAccountsList(false);
    LVAccounts.ItemIndex:= oldndx-1;
  end;
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
begin
  MnuRestore.Enabled:= (WindowState=wsMaximized) or (WindowState=wsMinimized);
  if MnuRestore.Enabled then MnuRestore.ImageIndex:=0 else MnuRestore.ImageIndex:=1 ;
  MnuMaximize.Enabled:= not (WindowState=wsMaximized);
  if MnuMaximize.Enabled then MnuMaximize.ImageIndex:=2 else MnuMaximize.ImageIndex:=3;
  MnuIconize.Enabled:= not (WindowState=wsMinimized);
  if MnuIconize.Enabled then MnuIconize.ImageIndex:= 4 else MnuIconize.ImageIndex:=5;
  if MnuGetAllMail.Enabled then MnuGetAllMail.ImageIndex:=6 else MnuGetAllMail.ImageIndex:=7;
  if MnuAbout.Enabled then MnuAbout.ImageIndex:=8 else MnuAbout.ImageIndex:=9;
  if MnuQuit.Enabled then MnuQuit.ImageIndex:=10 else MnuQuit.ImageIndex:=11;
end;

procedure TFMailsInBox.SGMailsBeforeSelection(Sender: TObject; aCol,
  aRow: Integer);
begin
  SGMails.Invalidate;
end;

procedure TFMailsInBox.SGMailsClick(Sender: TObject);
begin
  ShowMessage('Test');
end;

procedure TFMailsInBox.SGMailsDrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
var
  R: TRect;
  bmp: Tbitmap;
  bmppos: integer;
  ndx: integer;
begin
  if arow=0 then exit;
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then exit;
  // remove selection highlight if the control has not the focus
  if not SGHasFocus then
  begin
    SGMails.Canvas.Brush.Color := clWindow;
    SGMails.Canvas.FillRect (ARect);
    SGMails.Canvas.font.Color:= clDefault;  ;
  end;
  // Add mail image
  if acol=0 then
  begin
    bmppos:= 0;
    R.Left:= ARect.Left+2;
    R.Top:= ARect.Top+2;
    R.Right:=R.Left+18;
    R.Bottom:=R.Top+16;
    Bmp:= Tbitmap.Create;
    if FAccounts.Accounts.GetItem(ndx).Mails.GetItem(aRow-1).MessageNew then
      bmppos:= 1;
    if Pos ('multipart', FAccounts.Accounts.GetItem(ndx).Mails.GetItem(aRow-1).MessageContentType) >0 then
       bmppos:= bmppos+2;
    if FAccounts.Accounts.GetItem(ndx).Mails.GetItem(aRow-1).MessageToDelete then
       bmppos:= 4;
    CropBitmap(MailPictures.Bitmap, bmp, bmppos );
    SGMails.Canvas.StretchDraw(R, bmp);
    bmp.free;
    SGMails.Canvas.TextOut(ARect.Left+22,ARect.Top+3, SGMails.Cells[aCol, aRow]);
  end else
  begin
    SGMails.Canvas.TextOut(ARect.Left+2,ARect.Top+3, SGMails.Cells[aCol, aRow]);
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
  col1, row1: integer;
begin
  if Button = TMouseButton.mbRight then
  begin
    Col1:=0;
    Row1:=0;
    if SGMails.RowCount<2 then exit;
    if visible then SGMails.SetFocus;
    SGMails.MouseToCell(X, Y, Col1, Row1);
    pf := SGMails.ClientToScreen(Point(X, Y));
    if row1>0 then
    begin
      SGMails.Row:= (row1);
      // Do not use the grids PopupMenu property, it
      // prevents this event handler comletely.
      // Instead, activate the menu manually here.
      MnuMails.Popup(pf.X, pf.Y);
    end;
  end;
end;

procedure TFMailsInBox.TrayTimerTimer(Sender: TObject);
var
  s: string;
begin
  if not CheckingMail then
  begin
    // every 2 seconds
    ILTray.GetBitmap(TrayTimerTick div 2, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
    if TrayTimerTick div 2 <ILtray.count-1 then inc (TrayTimerTick, 2) else TrayTimerTick:= 0;
  end;
  // DSisplay date and time every second

  s:= DateTimetoStr(now, StatusFmtSets);
  s[1]:= upCase(s[1]);        // First letter of day in uppercase
  LNow.Caption:= s;
end;

procedure TFMailsInBox.BtnSettingsClick(Sender: TObject);
var
  i, oldlng, oldmailsel: integer;
begin
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
      if Settings.SmallBtns <> CBSmallBtns.Checked then  // Buttons size has changed
      SetSmallBtns(CBSmallBtns.Checked);
      Settings.Notifications:= CBNotifications.Checked;
      Settings.NoCloseAlert:= CBNoCloseAlert.Checked;
      Settings.NoQuitAlert:= CBNoCloseAlert.Checked;;
      Settings.SmallBtns:= CBSmallBtns.Checked;
      Settings.MailClient:= MailClients[CBMailClient.ItemIndex].Command;
      Settings.MailClientName:= MailClients[CBMailClient.ItemIndex].Name;
      Settings.MailClientIsUrl:= CBUrl.Checked;
      Settings.SoundFile:= ESoundFile.Text;
      Settings.LangStr := LangNums.Strings[CBLangue.ItemIndex];
      if FSettings.CBLangue.ItemIndex <> oldlng then ModLangue;
      if SettingsChanged then
      begin
        PopulateAccountsList(false);  // Needed to change language on hints
        LogAddLine(-1, now, sSettingsChange);
      end;
    end;
  end;
end;


procedure TFMailsInBox.BtnAboutClick(Sender: TObject);
begin
  // If main windows is hidden, place the about bopx at the center of desktop,
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

procedure TFMailsInBox.BtnDeleteClick(Sender: TObject);
begin
  ShowMessage('Todo: delete selected account');
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
//  CurAcc: TAccount;
  ndx: integer;
begin
  ndx:= LVAccounts.ItemIndex;
  if ndx<0 then exit;
  CheckingMail:= status;
//  CurAcc:= FAccounts.Accounts.GetItem(ndx);
  EnableControls(not status);
  if status then
  begin
    ChkMailTimerTick:= 0;
    ChkMailTimer.StartTimer;
    Screen.Cursor:= crHourGlass;
  end else
  begin
    ChkMailTimer.StopTimer;
    TrayTimerTick:= 0;
    ILTray.GetBitmap(TrayTimerTick, TrayTimerBmp);
    TrayMail.Icon.Assign(TrayTimerbmp);
    TrayMail.Hint:= sTrayHintNoMsg;
    Screen.Cursor:= DefCursor;
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
              ErrStr:= Format(ConnectErrorMsg, [E.Message]);
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
              ErrStr:= Format(ConnectErrorMsg, [E.Message]);
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
                  ErrStr:= Format(HeaderErrorMsg, [E.Message]);
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
                  ErrStr:= Format(HeaderErrorMsg, [E.Message]);
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
        GetMailInfos(CurName, Mail, IdMsg, siz);
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
          if length(CurAcc.UIDLToDel)>0 then            // delete messages with uidl in array
          begin
            for j:= 0 to length(CurAcc.UIDLToDel)-1 do
            begin
              if CurAcc.Mails.GetItem(i).MessageUIDL= CurAcc.UIDLToDel[j]  then
                 if IdPOP3_1.Delete(i) then
                 begin
                   Mails.Delete(i);
                   LogAddLine(CurAcc.UID, now, Format(sMsgDeleted, [CurAcc.Name, i+1]));
                end else LogAddLine(CurAcc.UID, now, Format(sMsgNotDeleted, [CurAcc.Name, i+1]));
            end;
          end;
        end;
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
              if IdIMAP4_1.UIDDeleteMsg(idMsg.UID) then
              begin
                Mails.Delete(i);
                 LogAddLine(CurAcc.UID, now, Format(sMsgDeleted, [CurAcc.Name, i+1]));
              end else LogAddLine(CurAcc.UID, now, Format(sMsgNotDeleted, [CurAcc.Name, i+1]));
            end;
          end;
        end;
        LStatus.Caption:= Format(sDisconnectServer, [Curacc.Name, CurAcc.Server]);
        LogAddLine(CurAcc.UID, now, LStatus.Caption );
        IdIMAP4_1.Disconnect(true);
      end;
    end;
  except
    on E: Exception do
    begin
      ErrStr:= Format(ConnectErrorMsg, [E.Message]);
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
  FAccounts.Accounts.ModifyField(index, 'LASTFIRE', now);
  Application.ProcessMessages;
  min:= EncodeTime(0,FAccounts.Accounts.GetItem(index).interval,0,0);
  FAccounts.Accounts.ModifyField(index, 'NEXTFIRE', now+min);
  FAccounts.Accounts.ModifyField(index, 'ERROR', Err);
  FAccounts.Accounts.ModifyField(index, 'ERRORSTR', ErrorsStr);
  TAccount(FAccounts.Accounts.Items[index]^).Mails.Reset;
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

// retrieve infos from mail header

procedure TFMailsInBox.GetMailInfos(CurName: String; var Mail: TMail; IdMsg: TIdMessage; siz: Integer);
var
  sfrom: string;
begin
  try
  sfrom:= IdMsg.From.Name;
  if length(sfrom)=0 then sfrom:= idMsg.From.Address;
  Mail.AccountName:= CurName;
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
// Currently : complete mailattente accounts
//             Outlook 2007-2013 accounts (password is not retrieved)

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
end;



// Log display, all system log or only account log

procedure TFMailsInBox.BtnLogClick(Sender: TObject);
var
  Curacc:TAccount;
  csvdoc: TCSVDocument;
  i: integer;
  s: string;
begin
  csvdoc:= TCSVDocument.Create;
  csvDoc.QuoteChar:='|';
  if FSettings.Settings.SaveLogs then csvdoc.CSVText:= MainLog+SessionLog.text
  else csvdoc.CSVText:= SessionLog.text;
   if TBitBtn(Sender).Name= 'BtnAccountLog' then
  begin
    if LVAccounts.ItemIndex>=0 then
    begin
      CurAcc:= FAccounts.Accounts.GetItem(LVAccounts.ItemIndex);
      s:='';
      for i:= 0 to csvdoc.RowCount-1 do
      begin
        // First insert *************** line to indicate a new session
        if pos('**', csvdoc.Cells[2,i]) >0 then s:= s+csvdoc.Cells[2,i]+#10;
        if csvdoc.Cells[0,i]= IntToStr(Curacc.UID) then
        s:=s+csvdoc.Cells[1,i]+' - '+csvdoc.Cells[2,i]+#10;
      end;
    end;
  end;
  if TBitBtn(Sender).Name= 'BtnLog' then
  begin
    for i:= 0 to csvdoc.RowCount-1 do
    begin
      if pos('**', csvdoc.Cells[2,i]) >0 then s:= s+csvdoc.Cells[2,i]+#10
      else s:=s+csvdoc.Cells[1,i]+' - '+csvdoc.Cells[2,i]+#10;
    end;
  end;
  With FLogView do
  begin
    Caption:= TBitBtn(Sender).Hint;
    RMLog.rtf:='';
    RMLog.Text:=s;
    RMLog.SelStart:=0;
    RMLog.Sellength:=0;
    showmodal;
  end;
  csvdoc.free;
end;

procedure TFMailsInBox.BtnQuitClick(Sender: TObject);
begin

  CanClose:= true;
  Close();
end;

// Disable controls during mail check to avoid conflicts
// Display hourglass cursor

procedure TFMailsInBox.EnableControls(Enable: boolean);
begin
  LVAccounts.Enabled:= enable;
  SGMails.Enabled:= enable;
  BtnImport.Enabled:= enable;
  BtnAccountLog.Enabled:= enable;
  BtnGetAllMail.Enabled:= enable;
  MnuGetAllMail.Enabled:= enable;
  MnuDeleteMsg.Enabled:= Enable;
  BtnGetAccMail.Enabled:= enable;
  BtnDelete.Enabled:= enable;
  BtnAddAcc.Enabled:= enable;
  BtnEditAcc.Enabled:= enable;
  BtnSettings.Enabled:= enable;
  BtnLog.Enabled:= enable;
  BtnAbout.Enabled:= enable;
  BtnQuit.Enabled:= enable;
  if enable then SCreen.Cursor:= DefCursor
  else Screen.Cursor:= crHourGlass;
end;

// Draw a circle with messages count on the account and tray icon

procedure TFMailsInBox.DrawTheIcon(Bmp: Tbitmap; NewCount: integer; CircleColor: TColor);
var
  i : integer;
  s: string;
  //Pict: TPicture;
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

// Animate tray icon during checking mail

procedure TFMailsInBox.OnChkMailTimer(Sender: TObject);
begin
  if CheckingMail then
  begin;
    TrayMail.Icon.LoadFromResourceName(HINSTANCE, 'XATT'+(InttoStr(ChkMailTimerTick)));
    inc (ChkMailTimerTick);
    if ChkMailTimerTick > 5 then ChkMailTimerTick:=0;
  end;
end;

// Load control captions and text variable translations
// from mailsinbox.lng

procedure TFMailsInBox.ModLangue;
begin
  with LangFile do
  begin
    LangStr:=FSettings.Settings.LangStr;
    sRetConfBack:= ReadString(LangStr,'RetConfBack','Recharge la dernière configuration sauvegardée');
    sCreNewConf:= ReadString(LangStr,'CreNewConf','Création d''une nouvelle configuration');
    sLoadConf:= ReadString(LangStr,'LoadConf','Chargement de la configuration');
    //Main Form
    Caption:=ReadString(LangStr,'Caption','Courrier en attente');
    OKBtn:= ReadString(LangStr, 'OKBtn','OK');
    YesBtn:=ReadString(LangStr,'YesBtn','Oui');
    NoBtn:=ReadString(LangStr,'NoBtn','Non');
    CancelBtn:=ReadString(LangStr,'CancelBtn','Annuler');
    BtnImport.Hint:=ReadString(LangStr,'BtnImport.Hint',BtnImport.Hint );
    BtnLogHint:=ReadString(LangStr,'BtnLogHint','Journal du compte %s');
    BtnGetAllMail.Hint:=ReadString(LangStr,'BtnGetAllMail.Hint',BtnGetAllMail.Hint);
    BtnGetAccMailHint:=ReadString(LangStr,'BtnGetAccMailHint','Vérifier le compte %s');
    BtnLaunchClient.Hint:=ReadString(LangStr,'BtnLaunchClient.Hint', BtnLaunchClient.Hint);
    BtnDeleteHint:=ReadString(LangStr,'BtnDeleteHint','Supprimer le compte %s');
    BtnAddAcc.Hint:=ReadString(LangStr,'BtnAddAcc.Hint',BtnAddAcc.Hint);
    BtnEditAccHint:=ReadString(LangStr,'BtnEditAccHint','Modifier le compte %s');
    BtnSettings.Hint:=ReadString(LangStr,'BtnSettings.Hint',BtnSettings.Hint);
    BtnAbout.Hint:=ReadString(LangStr,'BtnAbout.Hint',BtnAbout.Hint);
    BtnClose.Hint:=Format(ReadString(LangStr,'BtnClose.Hint',BtnClose.Hint),[#10]);
    BtnQuit.Hint:=ReadString(LangStr,'BtnQuit.Hint',BtnQuit.Hint);
    SGMails.Columns[0].Title.Caption:=ReadString(LangStr,'SGMails.Columns_0.Title.Caption',SGMails.Columns[0].Title.Caption);
    SGMails.Columns[1].Title.Caption:=ReadString(LangStr,'SGMails.Columns_1.Title.Caption',SGMails.Columns[1].Title.Caption);
    SGMails.Columns[2].Title.Caption:=ReadString(LangStr,'SGMails.Columns_2.Title.Caption',SGMails.Columns[2].Title.Caption);
    SGMails.Columns[3].Title.Caption:=ReadString(LangStr,'SGMails.Columns_3.Title.Caption',SGMails.Columns[3].Title.Caption);
    SGMails.Columns[4].Title.Caption:=ReadString(LangStr,'SGMails.Columns_4.Title.Caption',SGMails.Columns[4].Title.Caption);
    MnuDeleteMsg.Caption:=ReadString(LangStr,'MnuDeleteMsg.Caption',MnuDeleteMsg.Caption);
    MnuAnswerMsg.Caption:=ReadString(LangStr,'MnuAnswerMsg.Caption',MnuAnswerMsg.Caption);
    MnuInfos.Caption:=ReadString(LangStr,'MnuInfos.Caption',MnuInfos.Caption);
    AccountCaption:=ReadString(LangStr,'AccountCaption','Compte: %s');
    EmailCaption:=ReadString(LangStr,'EmailCaption','Courriel: %s');
    LastCheckCaption:=ReadString(LangStr,'LastCheckCaption','Dernière vérification: %s');
    NextCheckCaption:=ReadString(LangStr,'NextCheckCaption','Prochaine vérification: %s');
    sAccountImported:=ReadString(LangStr,'AccountImported','%d compte %s importé');
    sAccountsImported:=ReadString(LangStr,'AccountsImported','%d comptes %s importés');
    sMsgFound:=ReadString(LangStr,'MsgFound','%d message trouvé');
    sMsgsFound:=ReadString(LangStr,'MsgsFound','%d messages trouvés');
    LStatusCaption:=ReadString(LangStr,'LStatusCaption','%s sur le compte %s le %');
    sConnectToServer:=ReadString(LangStr,'ConnectToServer','%s : Connexion au serveur %s');
    sConnectedToServer:=ReadString(LangStr,'ConnectedToServer','%s : Connecté au serveur %s');
    ConnectErrorMsg:=ReadString(LangStr,'ConnectErrorMsg','Erreur de connexion : %s');
    HeaderErrorMsg:=ReadString(LangStr,'HeaderErrorMsg','Erreur d''obtention de l''entête : %s');
    sDisconnectServer:=ReadString(LangStr,'DisconnectServer','%s : Déonnexion du serveur %s');
    DisconnectedServer:=ReadString(LangStr,'DisconnectedServer','%s : Déconnecté du serveur %s');
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
    sNoQuitAlert:=ReadString(LangStr,'NoQuitAlert','Le programme va se fermer et ne vérifiera plus '+
                                     'l''arrivée de nouveaux courriels. Pour que la vérification '+
                                     'se poursuive en tâche de fond, cliquez sur le bouton "Quitter Courrier en attente".');
    sNoCloseAlert:=ReadString(LangStr,'NoCloseAlert','Le programme va se poursuivre en tâche de fond '+
                                     'pour vérifier l''arrivée de nouveaux courriels. Pour quitter le '+
                                     'programme, cliquez sur le bouton "Fermer la fenêtre de Courrier en attente".');
    sCannotQuit:=ReadString(LangStr,'CannotQuit','Impossible de quitter pendant la vérification de courriels');
    sClosingProg:=ReadString(LangStr,'ClosingProg','Fermeture de Courriels en attente');
    sRestart:=ReadString(LangStr,'Restart','Redémarrage après arrêt forcé');
    sSettingsChange:=ReadString(LangStr,'SettingsChange','Configuration modifiée');
    sMsgDeleted:=ReadString(LangStr,'MsgDeleted','%s : Message %u supprimé');
    sMsgNotDeleted:=ReadString(LangStr,'MsgNotDeleted','%s : Message %u non supprimé');
    sMnuDelMsg:=ReadString(LangStr,'MnuDelMsg', 'Effacer le courriel "%s"');
    sMnuAnswerMsg:=ReadString(LangStr,'MnuAnswerMsg','Répondre au courriel "%s"');
    sAlertDelMmsg:=ReadString(LangStr,'AlertDelMmsg','Voulez-vous supprimer le courriel "%s" ?');

    // About
    sNoLongerChkUpdates:=ReadString(LangStr,'NoLongerChkUpdates','Ne plus rechercher les mises à jour');
    sLastUpdateSearch:=ReadString(LangStr,'LastUpdateSearch','Dernière recherche de mise à jour');
    sUpdateAvailable:=ReadString(LangStr,'UpdateAvailable','Nouvelle version %s disponible');
    sUpdateAlertBox:=ReadString(LangStr,'UpdateAlertBox','Version actuelle: %sUne nouvelle version %s est disponible');
    Aboutbox.Caption:=ReadString(LangStr,'Aboutbox.Caption','A propos du Gestionnaire de Contacts');
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
    FAccounts.LMailClient.Caption:=ReadString(LangStr,'FAccounts.LMailClient.Caption',FAccounts.LMailClient.Caption);
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
    FAccounts.BtnMailClient.Hint:=ReadString(LangStr,'FAccounts.BtnMailClient.Hint',FAccounts.BtnMailClient.Hint);
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
    FSettings.LMailClient.Caption:=FAccounts.LMailClient.Caption;
    FSettings.LSoundFile.Caption:=FAccounts.LSoundFile.Caption;
    FSettings.LLangue.Caption:=ReadString(LangStr,'FSettings.LLangue.Caption',FSettings.LLangue.Caption);
    FSettings.BtnMailClient.Hint:=FAccounts.BtnMailClient.Hint;
    FSettings.BtnPlaySound.Hint:=FAccounts.BtnPlaySound.Hint;
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
    FMailClientChoose.BtnMailClient.Hint:=FAccounts.BtnMailClient.Hint;

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
    MnuQuit.Caption:= BtnQuit.Hint;
    MnuAbout.Caption:=BtnAbout.Hint;

    // Alertbox
    AlertBox.BtnCancel.Caption:= CancelBtn;
    AlertBox.BtnOK.Caption:= OKBtn;
    AlertBox.CBNoShowAlert.Caption:=ReadString(LangStr,'AlertBox.CBNoShowAlert.Caption',AlertBox.CBNoShowAlert.Caption);
  end;

end;

end.

