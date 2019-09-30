unit peUtils;

{$mode Delphi}

interface

uses
  Classes, SysUtils, IniFiles;

const
  MaxOleEnumIndex = 1000;
  msoPicture = $0000000D;
  msoScaleFromTopLeft = $00000000;
  msoPictureAutomatic = $00000001;

  RO = 1;   // Read-Only
  RW = 0;   // Read/Write
  TRI = 2;  // undefiniert

  // Image indexes
  IMG_INDEX_POWERPOINT = 0;
  IMG_INDEX_SHAPE = 1;
  IMG_INDEX_SLIDE = 3;
  IMG_INDEX_HEADER_FOOTER = 4;
  IMG_INDEX_SLIDE_NUMBER = 5;
  IMG_INDEX_DATE_TIME = 6;
  IMG_INDEX_HEADER = 7;
  IMG_INDEX_FOOTER = 8;
  IMG_INDEX_PROPERTIES = 9;
  IMG_INDEX_OLE_FORMAT = 10;
  IMG_INDEX_PICTURE_FORMAT = 11;
  IMG_INDEX_TEXT_FRAME = 12;
  IMG_INDEX_TEXT_EFFECT = 13;
  IMG_INDEX_SHADOW = 14;
  IMG_INDEX_ANIMATION_SETTINGS = 15;
  IMG_INDEX_SLIDESHOW_TRANSITION = 16;
  IMG_INDEX_SLIDE_SHOW_SETTINGS = 17;
  IMG_INDEX_PAGE_SETUP = 18;
  IMG_INDEX_MASTER = 19;
  IMG_INDEX_FILL = 20;
  IMG_INDEX_READWRITE = 21;
  IMG_INDEX_READONLY = 22;
  IMG_INDEX_ACTION_SETTINGS = 23;
  IMG_INDEX_CALLOUT = 24;
  IMG_INDEX_DIAGRAM = 25;
  IMG_INDEX_ERROR = 26;
  IMG_INDEX_EFFECT = 27;
  IMG_INDEX_SEQUENCE = 28;
  IMG_INDEX_EFFECT_INFORMATION = 29;
  IMG_INDEX_STAR = 30;
  IMG_INDEX_LINK = 31;

type
  TOleEnumItem = record
    Name : string;
    Value : longword;
  end;

  TOleEnumArray = array[0..MaxOleEnumIndex] of TOleEnumItem;

function GetPowerpointVersion(PPTApp: OLEVariant) : integer;
procedure OriginalPictureSize(AShape: OLEVariant; out AWidth, AHeight:integer);

function PixelsToPoints(Pixels: Integer): Single;
function PointsToPixels(Points: double): Integer;
function PointsToCm(Points: double) : double;

function IsRelativePath(const AFileName: string): boolean;

function CreateIni: TCustomIniFile;
function GetScreenWorkAreaRect: TRect;


implementation

uses
  Graphics, LCLIntf, LCLType, Forms;

const
  INCH = 2.54;     // 1" = 2.54 cm

function GetPowerpointVersion(PptApp: OLEVariant): integer;
var
  ver: string;
begin
  ver := PptApp.Version;
  ver := copy(ver, 1, pos('.', ver) - 1);
  result := StrToInt(ver);
end;

procedure ResetPictureProperties(AShape: OLEVariant);
// from: http://www.mvps.org/skp/ppt00044.htm#2
begin
  if AShape.&Type = msoPicture then
  begin
    AShape.LockAspectRatio := false;
    AShape.ScaleWidth(1, true, msoScaleFromTopLeft);
    AShape.ScaleHeight(1, true, msoScaleFromTopLeft);
    AShape.LockAspectRatio := true;
    AShape.PictureFormat.ColorType := msoPictureAutomatic;
    AShape.PictureFormat.Brightness := 0.5;
    AShape.PictureFormat.Contrast := 0.5;
  end;
end;

procedure OriginalPictureSize(AShape: OLEVariant; out AWidth, AHeight:integer);
var
  sh: OLEVariant;
begin
  if AShape.&Type = msoPicture then
  begin
    sh := AShape.Duplicate;
    ResetPictureProperties(sh);
    AWidth := PointsToPixels(sh.Width);
    AHeight := PointsToPixels(sh.Height);
    sh.Delete;
  end else begin
    AWidth := -1;
    AHeight := -1;
  end;
end;

function PixelsPerInch: integer; inline;
begin
  Result := ScreenInfo.PixelsPerInchX;
end;

function PixelsToPoints(Pixels: Integer): Single;
begin
  Result := (Pixels / PixelsPerInch) * 72;
end;

function PointsToPixels(Points: double): Integer;
begin
  Result := Trunc((Points / 72) * PixelsPerInch);
end;

function PointsToCm(Points: double) : double;
begin
  Result := INCH * Points / 72.0;
end;

function IsRelativePath(const AFileName: string): boolean;
begin
  Result := (Length(AFileName) > 1) and
    (not (AFileName[1] in [PathDelim, '~', '$'])) and
    (AFileName[2] <> DriveDelim);
end;

function CreateIni: TCustomIniFile;
begin
  Result := TMemIniFile.Create('./pptex.ini');
end;

function GetScreenWorkAreaRect: TRect;
begin
  Result := Screen.WorkAreaRect;
  if (Result.Left = 0) and (Result.Right = Screen.Width) and (Result.Top = 0) and (Result.Bottom = Screen.Height) then
    // no Taskbar
  else
  begin
    Result.Left -= GetSystemMetrics(SM_CXSIZEFRAME);
    Result.Bottom -= GetSystemMetrics(SM_CYSIZEFRAME) + GetSystemMetrics(SM_CYCAPTION);
  end;
end;

end.

