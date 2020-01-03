unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, BaseUnix, process;

type

  { TForm1 }

  TForm1 = class(TForm)
    procedure FormActivate(Sender: TObject);
  private
    function parseMimeApps(filename: string): string;
  public
    AppListArr : array of string;
    mailapp: string;

  end;



var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

{[Desktop Entry]
Type=Application
Encoding=UTF-8
Name=Sample Application Name
Comment=A sample application
Exec=application
Icon=application.png
Terminal=false  }

function TForm1.parseMimeApps(filename: string): string;
var
  sl: TstringList;
  i: integer;
  p: integer;
begin
  result:= '';
  sl:= TstringList.Create;
  sl.LoadFromFile(filename);
   //x-scheme-handler/mailto
  for i:= 0 to sl.Count-1 do
  begin
    if pos('x-scheme-handler/mailto=', sl.Strings[i])=1 then
    begin
     result:= Copy(sl.Strings[i], 25, length(sl.Strings[i])-25);
     break;
    end;
  end;
  sl.free;
end;

procedure TForm1.FormActivate(Sender: TObject);
var
  HomeDir: string;
  MimeAppsFile: string;
  DsktopFile: string;
  username: string;
  StrDskFile: TstringList;
  i: integer;
  P: TProcess;
      Function ExecParam(Param: String): String;
        Begin
          P.Parameters[0]:= Param;
          P.Execute;
          SetLength(Result, 1000);
          SetLength(Result, P.Output.Read(Result[1], Length(Result)));
          While (Length(Result) > 0) And (Result[Length(Result)] In [#8..#13,#32]) Do
            SetLength(Result, Length(Result) - 1);
        End;
begin
  HomeDir:= GetUserDir;
  Setlength(AppListArr, 9);
  AppListArr[0]:= HomeDir+'.local/share/applications/mimeapps.list';
  AppListArr[1]:= HomeDir+'.local/share/applications/mimeinfo.cache';
  AppListArr[2]:= HomeDir+'.local/share/applications/defaults.list';
  AppListArr[3]:= '/usr/local/share/applications/mimeapps.list';
  AppListArr[4]:= '/usr/local/share/applications/mimeinfo.cache';
  AppListArr[5]:= '/usr/local/share/applications/defaults.list';
  AppListArr[6]:= '/usr/share/applications/mimeapps.list';
  AppListArr[7]:= '/usr/share/applications/mimeinfo.cache';
  AppListArr[8]:= '/usr/share/applications/defaults.list';
  For i:= 0 to 8 do
  begin
    if FileExists(AppListArr[i]) then
    begin
      DsktopFile:= parseMimeApps(AppListArr[i]);
      if length(DsktopFile) > 0 then break;
    end;
  end;
  //ShowMessage(DsktopFile);
  P:= TProcess.Create(Nil);
  P.Options:= [poWaitOnExit, poUsePipes];
  P.Executable:= 'thunderbird';
  P.Execute;

///usr/local/share/applications/mimeapps.list
///usr/local/share/applications/defaults.list
///usr/share/applications/mimeapps.list
///usr/share/applications/defaults.list
  {if not DirectoryExists(HomeDir+'.config/autostart/') then CreateDir(HomeDir+'.config/autostart/');
  If not FileExists('contactmgr.desktop') then
  begin
    StrDskFile:= TstringList.Create;
    StrDskFile.Add('[Desktop Entry]');
    StrDskFile.Add('Type=Application');
    StrDskFile.Add('Encoding=UTF-8');
    StrDskFile.Add('Name=ContactMgr');
    StrDskFile.Add('Comment=Contacts manager');
    StrDskFile.Add('Exec=/home/bernard/Documents/lazarus/foxbirdbackup/foxbirdbackuplinux');
    StrDskFile.Add('Terminal=false');
    StrDskFile.SaveToFile(HomeDir+'.config/autostart/contactmgr.desktop');
    fpChmod (HomeDir+'.config/autostart/contactmgr.desktop',&777);
    end;    }
    //HomeShareDir:= GetEnvironmentVariable('$XDG_DATA_HOME');


    //if DirectoryExists(HomeShareDir) then
   // ShowMessage(HomeShareDir);
//If FileExists(HomeDir+'.config/autostart/contactmgr.desktop') then DeleteFile(HomeDir+'.config/autostart/contactmgr.desktop');
   end;

end.
