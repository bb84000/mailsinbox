{*******************************************************************************}
{ mailclient : Unit to find mail client(s)                                      }
{ for mailsinbox application                                                    }
{ bb -sdtp - january 2023                                                       }
{*******************************************************************************}

unit mailclients1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Buttons, ExtCtrls,
  StdCtrls, registry, lazbbinifiles;

type
  // Define the classes in this Unit at the very start for clarity
  TFMailClientChoose = Class;          // This is a forward class definition

  { TFMailClientChoose }

  TFMailClientChoose = class(TForm)
    BtnMailClient: TSpeedButton;
    BtnOk: TBitBtn;
    BtnCancel: TBitBtn;
    CBUrl: TCheckBox;
    EName: TEdit;
    ECommand: TEdit;
    LName: TLabel;
    LCommand: TLabel;
    OD1: TOpenDialog;
    PnlButtons: TPanel;
    procedure BtnMailClientClick(Sender: TObject);
    procedure CBUrlChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private

  public
    procedure Translate(LngFile: TBbIniFile);
  end;

var
  FMailClientChoose: TFMailClientChoose;

implementation

{$R *.lfm}

{ TFMailClientChoose }

procedure TFMailClientChoose.CBUrlChange(Sender: TObject);
begin
  BtnMailClient.Enabled:= not CBUrl.checked;
end;

procedure TFMailClientChoose.FormActivate(Sender: TObject);
begin
  // Center buttons in case of width change
  BtnOK.Left:= (PnlButtons.ClientWidth-BtnOK.width*2-20) div 2;
  BtnCancel.Left:= BtnOK.Left+BtnOK.Width+20;
end;

procedure TFMailClientChoose.BtnMailClientClick(Sender: TObject);
begin
  if OD1.Execute then
  ECommand.text:= OD1.FileName;
end;

procedure TFMailClientChoose.Translate(LngFile: TBbIniFile);
begin
  if assigned (Lngfile) then
  with LngFile do
  begin
    BtnOK.Caption:= ReadString('Common', 'OKBtn', BtnOK.Caption);
    BtnCancel.Caption:= ReadString('Common', 'CancelBtn', BtnCancel.Caption);
    Caption:=ReadString('FMailClientChoose', 'Caption', Caption);
    LName.Caption:=ReadString('FMailClientChoose', 'LName.Caption', LName.Caption);
    LCommand.Caption:=ReadString('FMailClientChoose', 'LCommand.Caption', LCommand.Caption);
    CBUrl.Hint:=ReadString('FMailClientChoose', 'CBUrl.Hint', CBUrl.Hint);
    BtnMailClient.Hint:=ReadString('FMailClientChoose', 'BtnMailClient.Hint', BtnMailClient.Hint);
  end;
end;

end.

