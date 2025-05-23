program mailsinbox;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, indylaz, accounts1, settings1, lazbbupdatedlg,
  impex1, mailclients1, log1, mailsinbox1, lazbbcomponents;

{$R *.res}
{$R mailinboxres.rc}

begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TFMailsInBox, FMailsInBox);
  Application.CreateForm(TFSettings, FSettings);
  Application.CreateForm(TFAccounts, FAccounts);
   Application.CreateForm(TFImpex, FImpex);
  Application.CreateForm(TFMailClientChoose, FMailClientChoose);
  Application.CreateForm(TFLogView, FLogView);
  Application.Run;
end.

