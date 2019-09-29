unit peAbout;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ButtonPanel, ExtCtrls,
  StdCtrls, Buttons;

type

  { TAboutForm }

  TAboutForm = class(TForm)
    BitBtn1: TBitBtn;
    Image1: TImage;
    lblIcons: TLabel;
    lblAppNameLong: TLabel;
    lblAcknowledgements: TLabel;
    lblCompiler: TLabel;
    lblFPC: TLabel;
    lblLazarus: TLabel;
    lblIDE: TLabel;
    lblAppName: TLabel;
    lblIcons8: TLabel;
    Panel1: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure lblURLClick(Sender: TObject);
    procedure lblURLMouseEnter(Sender: TObject);
    procedure lblURLMouseLeave(Sender: TObject);
  private

  public

  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.lfm}

uses
  LCLIntf, Types;
//  peGlobal;

{ TAboutForm }

procedure TAboutForm.FormCreate(Sender: TObject);
var
  lSize: TSize;
begin
  Image1.Picture.Assign(Application.Icon);
  lSize := Size(Image1.Width, Image1.Width);
  Image1.Picture.Icon.Current := Image1.Picture.Icon.GetBestIndexForSize(lSize);
end;

procedure TAboutForm.lblURLClick(Sender: TObject);
begin
  if Sender = lblFPC then
    OpenURL('https://www.freepascal.org/')
  else if Sender = lblLazarus then
    OpenURL('https://www.lazarus-ide.org/')
  else if Sender = lblIcons8 then
    OpenURL('http://www.icons8.com');
end;

procedure TAboutForm.lblURLMouseEnter(Sender: TObject);
begin
  with (Sender as TControl).Font do
    Style := Style + [fsUnderline];
end;

procedure TAboutForm.lblURLMouseLeave(Sender: TObject);
begin
  with (Sender as TControl).Font do
    Style := Style - [fsUnderline];
end;

end.

