unit peUtils;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  MaxOleEnumIndex = 1000;
  msoPicture = $0000000D;
  msoScaleFromTopLeft = $00000000;
  msoPictureAutomatic = $00000001;

  RO = 1;   // Read-Only
  RW = 0;   // Read/Write
  TRI = 2;  // undefiniert

  // Image indexes
  IMG_INDEX_READWRITE = 21;
  IMG_INDEX_READONLY = 22;
  IMG_INDEX_ERROR = 26;

type
  TOleEnumItem = record
    Name : string;
    Value : longword;
  end;

  TOleEnumArray = array[0..MaxOleEnumIndex] of TOleEnumItem;

function  GetPowerpointVersion(PPTApp: OLEVariant) : integer;

function  PixelsToPoints(Pixels: Integer): Single;
function  PointsToPixels(Points: double): Integer;
function  PointsToCm(Points: double) : double;

function IsRelativePath(const AFileName: string): boolean;


implementation

uses
  Graphics;

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

end.

