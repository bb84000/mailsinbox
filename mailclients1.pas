{*******************************************************************************}
{ mailclient : Unit to find mail client(s)                                      }
{ for mailsinbox application                                                    }
{ bb -sdtp - january 2020                                                       }
{*******************************************************************************}

unit mailclients1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Buttons, ExtCtrls,
  StdCtrls, registry;

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

end.

