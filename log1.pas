{*******************************************************************************}
{ log1 : Unit to display program log                                            }
{ for mailsinbox application                                                    }
{ bb -sdtp - january 2023                                                       }
{*******************************************************************************}

unit log1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons,
  Menus, RichMemo, Clipbrd, lazutf8, lazbbutils, lazbbinifiles;

type

  { TFLogView }

  TFLogView = class(TForm)
    BtnOK: TBitBtn;
    ILMnuLog: TImageList;
    MnuCopyLine: TMenuItem;
    MnuCopyAll: TMenuItem;
    MnuCopySel: TMenuItem;
    PnlButtons: TPanel;
    MnuCopy: TPopupMenu;
    RMLog: TRichMemo;
    procedure FormChangeBounds(Sender: TObject);
    procedure MnuCopyAllClick(Sender: TObject);
    procedure MnuCopyLineClick(Sender: TObject);
    procedure MnuCopyPopup(Sender: TObject);
    procedure MnuCopySelClick(Sender: TObject);
    procedure RMLogChange(Sender: TObject);
  private

  public
    procedure Translate(LngFile: TBbIniFile);
  end;

var
  FLogView: TFLogView;

implementation

{$R *.lfm}

uses mailsinbox1;

{ TFLogView }

// copy selection

procedure TFLogView.MnuCopySelClick(Sender: TObject);
begin
  Clipboard.AsText := RMLog.SelText;
end;

procedure TFLogView.RMLogChange(Sender: TObject);
begin

end;

// copy line

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

procedure TFLogView.MnuCopyPopup(Sender: TObject);
var
  bmp: TBitmap;
begin
  bmp:= Tbitmap.Create;
  ILMnuLog.GetBitmap(0, Bmp);
  MnuCopySel.enabled:=  RMLog.SelLength> 0;
  CropBitmap(bmp, MnuCopySel.Bitmap, MnuCopySel.enabled);
  ILMnuLog.GetBitmap(1, Bmp);
  CropBitmap(bmp, MnuCopyAll.Bitmap, MnuCopyAll.enabled);
  ILMnuLog.GetBitmap(2, Bmp);
  MnuCopyLine.enabled:=  (RMLog.SelStart<RMlog.GetTextLen);
  CropBitmap(bmp, MnuCopyLine.Bitmap, MnuCopyLine.enabled);
  if Assigned(bmp) then bmp.free;
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

procedure TFLogView.Translate(LngFile: TBbIniFile);
begin
  if assigned (Lngfile) then
  with LngFile do
  begin
    BtnOK.Caption:= ReadString('Common', 'OKBtn', BtnOK.Caption);
    MnuCopySel.Caption:= ReadString('FLogView', 'MnuCopySel.Caption', MnuCopySel.Caption);
    MnuCopyAll.Caption:= ReadString('FLogView', 'MnuCopyAll.Caption', MnuCopyAll.Caption);
    MnuCopyLine.Caption:= ReadString('FLogView', 'MnuCopyLine.Caption', MnuCopyLine.Caption);
  end;

end;




end.

