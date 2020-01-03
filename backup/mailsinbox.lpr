program mailsinbox;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, indylaz, mailsinbox1, accounts1, settings1, lazbbabout, lazbbalert,
  impex1, mailclients1
  { you can add units after this };

{$R *.res}
{$R mailinboxres.rc}

begin
  RequireDerivedFormResource:=True;
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TFMailsInBox, FMailsInBox);
  Application.CreateForm(TFSettings, FSettings);
  Application.CreateForm(TFAccounts, FAccounts);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TAlertBox, AlertBox);
  Application.CreateForm(TFImpex, FImpex);
  Application.CreateForm(TFMailClients, FMailClients);
  Application.Run;
end.

