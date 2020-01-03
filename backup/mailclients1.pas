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
    procedure CBUrlChange(Sender: TObject);
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

end.

