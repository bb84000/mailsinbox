unit log1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  Menus, RichMemo, Clipbrd, lazutf8;

type

  { TFLogView }

  TFLogView = class(TForm)
    BtnOK: TBitBtn;
    MnuCopyLine: TMenuItem;
    MnuCopyAll: TMenuItem;
    MnuCopySel: TMenuItem;
    PnlButtons: TPanel;
    MnuCopy: TPopupMenu;
    RMLog: TRichMemo;
    procedure FormChangeBounds(Sender: TObject);
    procedure MnuCopyAllClick(Sender: TObject);
    procedure MnuCopyLineClick(Sender: TObject);
    procedure MnuCopySelClick(Sender: TObject);
  private

  public

  end;

var
  FLogView: TFLogView;

implementation

{$R *.lfm}

uses mailsinbox1;

{ TFLogView }


procedure TFLogView.MnuCopySelClick(Sender: TObject);
begin
  Clipboard.AsText := RMLog.SelText;

end;


procedure TFLogView.MnuCopyLineClick(Sender: TObject);
var
  selbeg: integer;
  linebeg, linelength: integer;
  i: integer;
begin
  selbeg:= RMLog.SelStart;
  linebeg:= 0;
  for i:=0 to RMLog.Lines.Count-1 do
  begin
    linelength:= UTF8length(RMLog.lines[i]);
    if (selbeg > linebeg) and (selbeg <= linebeg+linelength) then
    begin
      RMLog.SelStart:= linebeg;
      RMLog.SelLength:= linelength;
      Clipboard.AsText := RMLog.SelText;
      break;
    end;
    linebeg:= linebeg+linelength+1;
  end;
end;



procedure TFLogView.MnuCopyAllClick(Sender: TObject);
begin
  RMLog.SelStart:=0;
  RMlog.SelLength:= length(RMLog.text);
  Clipboard.AsText:= RMLog.Text;
end;

procedure TFLogView.FormChangeBounds(Sender: TObject);
begin
  FMailsInBox.OnChangeBounds(sender);
end;







end.

