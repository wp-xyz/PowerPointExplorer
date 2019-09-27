unit pePropertyViewer;

interface

uses
  //Windows, Messages,
  SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls;

type

  { TPropertyForm }

  TPropertyForm = class(TForm)
    Panel1: TPanel;
    btnOK: TBitBtn;
    lblParamName: TLabel;
    Panel2: TPanel;
    ParamMemo: TMemo;
    ColorViewer: TShape;
    procedure FormShow(Sender: TObject);
  private
    { Private-Deklarationen }
    FDisplayOption: string;
    function  GetPropertyName: string;
    function  GetPropertyValue: string;
    procedure SetDisplayOption(const AValue: string);
    procedure SetPropertyName(const AValue: string);
    procedure SetPropertyValue(const AValue: string);
  public
    { Public-Deklarationen }
    function ParseColor: TColor;
    property DisplayOption: string read FDisplayOption write SetDisplayOption;
    property PropertyName: string read GetPropertyName write SetPropertyName;
    property PropertyValue: string read GetPropertyvalue write SetPropertyValue;
  end;

var
  PropertyForm: TPropertyForm;

implementation

{$R *.lfm}

uses
  StrUtils;

function TPropertyForm.GetPropertyName: string;
begin
  Result := lblParamName.Caption;
  Delete(result, Length(result), 1);
end;

function TPropertyForm.GetPropertyValue: string;
begin
  Result := ParamMemo.Lines.Text;
end;

procedure TPropertyForm.SetDisplayOption(const AValue: string);
begin
  if AValue <> FDisplayOption then
  begin
    FDisplayOption := AValue;
    ColorViewer.Visible := (FDisplayOption <> '') and (FDisplayOption[1] = 'C');
  end;
end;

procedure TPropertyForm.SetPropertyName(const AValue: string);
begin
  lblParamName.Caption := AValue + ':';
end;

procedure TPropertyform.SetPropertyValue(const AValue: string);
begin
  ParamMemo.Lines.Text := StringReplace(AValue, #$b, LineEnding, [rfReplaceAll]);
end;

function TPropertyForm.ParseColor: TColor;
var
  p, q: integer;
  txt: String;
begin
  txt := ParamMemo.Lines.Text;
  p := pos('$', txt);
  q := Pos(' ', txt);
  if q < p then q := Pos(')', txt);
  txt := Copy(txt, p, q-p);
  Result := TColor(StrToInt(txt));
end;

procedure TPropertyForm.FormShow(Sender: TObject);
begin
  if (FDisplayOption <> '') then
  begin
    case FDisplayOption[1] of
      'C' : ColorViewer.Brush.Color := ParseColor;
    end;
  end;
end;

end.
