unit peOffice;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  MaxOleEnumIndex = 1000;

type
  TOleEnumItem = record
    Name: string;
    Value: longword;
  end;

  TOleEnumArray = array[0..MaxOleEnumIndex] of TOleEnumItem;

function GetOleEnumName(const Enums; Count: Integer; Value: LongWord): string;
function GetOleEnumText(const Enums; Count: Integer; value: LongWord): string;

const
  // Constants for enum MsoTriState
  msoTrue = $FFFFFFFF;
  msoFalse = $00000000;
  msoCTrue = $00000001;
  msoTriStateToggle = $FFFFFFFD;
  msoTriStateMixed = $FFFFFFFE;

  // copied from the type libraries
  // msoShapetype
  msoShapeTypeMixed = $FFFFFFFE;
  msoAutoShape = $00000001;
  msoCallout = $00000002;
  msoChart = $00000003;
  msoComment = $00000004;
  msoFreeform = $00000005;
  msoGroup = $00000006;
  msoEmbeddedOLEObject = $00000007;
  msoFormControl = $00000008;
  msoLine = $00000009;
  msoLinkedOLEObject = $0000000A;
  msoLinkedPicture = $0000000B;
  msoOLEControlObject = $0000000C;
  msoPicture = $0000000D;
  msoPlaceholder = $0000000E;
  msoTextEffect = $0000000F;
  msoMedia = $00000010;
  msoTextBox = $00000011;
  msoScriptAnchor = $00000012;
  msoTable = $00000013;
  msoScaleFromTopLeft = $00000000;
  msoPictureAutomatic = $00000001;

  ppAdvanceModeMixed = $FFFFFFFE;
  ppAdvanceOnClick = $00000001;
  ppAdvanceOnTime = $00000002;
  ppLayoutBlank = $0000000C;
  ppViewSlide = $00000001;
  ppViewSlideSorter = $00000007;
  ppViewSlideMaster = $00000002;
  ppViewNotesPage = $00000003;
  ppViewHandoutMaster = $00000004;
  ppViewNotesMaster = $00000005;
  ppViewTitleMaster = $00000008;
  ppSoundEffectsMixed = $FFFFFFFE;
  ppSoundNone = $00000000;
  ppSoundStopPrevious = $00000001;
  ppSoundFile = $00000002;

  ppMouseClick = $00000001;
  ppMouseOver = $00000002;

function MsoAlignCmdStr(AValue: LongWord): string;
function MsoAnimAccumulateStr(AValue: LongWord): string;
function MsoAnimAdditiveStr(AValue: LongWord): string;
function MsoAnimAfterEffectStr(AValue: LongWord): string;
function MsoAnimateByLevelStr(AValue: Longword): string;
function MsoAnimationTypeStr(AValue: LongWord): string;
function MsoAnimDirectionStr(AValue: LongWord): string;
function MsoAnimEffectAfterStr(AValue: LongWord): string;
function MsoAnimEffectRestartStr(AValue: Longword): string;
function MsoAnimEffectStr(AValue: LongWord): string;
function MsoAnimPropertyStr(AValue: LongWord): string;
function MsoAnimTextUnitEffectStr(AValue: LongWord): string;
function MsoAnimTriggerTypeStr(AValue: LongWord): string;
function MsoAnimTypeStr(AValue: LongWord): string;
function MsoAppLanguageIDStr(AValue: LongWord): string;
function MsoArrowHeadLengthStr(AValue: LongWord): string;
function MsoArrowHeadStyleStr(AValue: LongWord): string;
function MsoArrowHeadWidthStr(AValue: LongWord): string;
function MsoAutoShapeTypeStr(AValue: LongWord): string;
function MsoBlackWhiteModeStr(AValue: LongWord): string;
function MsoCalloutAngleTypeStr(AValue: LongWord): string;
function MsoCalloutDropTypeStr(AValue: LongWord): string;
function MsoCalloutTypeStr(AValue: LongWord): string;
function MsoColorTypeStr(AValue: LongWord): string;
function MsoConnectorTypeStr(AValue: LongWord): string;
function MsoDiagramTypeStr(AValue: LongWord): string;
function MsoDistributeCmdStr(AValue: LongWord): string;
function MsoEditingTypeStr(AValue: LongWord): string;
function MsoFillTypeStr(AValue: LongWord): string;
function MsoGradientColorTypeStr(AValue: LongWord): string;
function MsoGradientStyleStr(AValue: LongWord): string;
function MsoHorizontalAnchorStr(AValue: LongWord): string;
function MsoLineDashStyleStr(AValue: LongWord): string;
function MsoLineStyleStr(AValue: LongWord): string;
function MsoOrientationStr(AValue: LongWord): string;
function MsoPatternTypeStr(AValue: LongWord): string;
function MsoPictureColorTypeStr(AValue: LongWord): string;
function MsoPresetGradientTypeStr(AValue: LongWord): string;
function MsoPresetLightingDirectionStr(AValue: LongWord): string;
function MsoPresetLightingSoftnessStr(AValue: LongWord): string;
function MsoPresetMaterialStr(AValue: LongWord): string;
function MsoPresetTextureStr(AValue: LongWord): string;
function MsoPresetThreeDFormatStr(AValue: LongWord): string;
function MsoScaleFromStr(AValue: LongWord): string;
function MsoSegmentTypeStr(AValue: LongWord): string;
function MsoShadowTypeStr(AValue: LongWord): string;
function MsoShapeTypeStr(AValue: Longword): string;
function MsoTextEffectAlignmentStr(AValue: LongWord): string;
function MsoTextEffectShapeStr(AValue: LongWord): string;
function MsoTextOrientationStr(AValue: LongWord): string;
function MsoTextureTypeStr(AValue: LongWord): string;
function MsoTriStateStr(AValue: LongWord): string;
function MsoVerticalAnchorStr(AValue: LongWord): string;
function MsoZOrderCmdStr(AValue: LongWord): string;

function PpActionTypeStr(AValue: LongWord): string;
function PpAdvanceModeStr(AValue: LongWord): string;
function PpAfterEffectStr(AValue: Longword): string;
function PpAutoSizeStr(AValue: LongWord): string;
function PpChartUnitEffectStr(AValue: LongWord): string;
function PpColorSchemeIndexStr(AValue: LongWord): string;
function PpDateTimeFormatStr(AValue: LongWord): string;
function PpEntryEffectStr(AValue: LongWord): string;
function PpFollowColorsStr(AValue: LongWord): string;
function PpMediaTypeStr(AValue: LongWord): string;
function PpPlaceholderTypeStr(AValue: LongWord): string;
function PpSlideLayoutStr(AValue: LongWord): string;
function PpSlideShowRangeTypeStr(AValue: LongWord): string;
function PpSlideShowTypeStr(AValue: LongWord): string;
function PpSlideSizeTypeStr(AValue: LongWord): string;
function PpSoundEffectTypeStr(AValue: LongWord): string;
function PpSoundFormatTypeStr(AValue: LongWord): string;
function PpTextLevelEffectStr(AValue: LongWord): string;
function PpTextUnitEffectStr(AValue: LongWord): string;
function PpUpdateOptionStr(AValue: LongWord): string;


implementation

function GetOleEnumName(const Enums; Count:integer; Value:longword) : string;
var
  i : integer;
begin
  if (Count<0) or (Count>=MaxOleEnumIndex)
    then raise Exception.Create('GetOleEnumName: Unzulässiger Index.');
  for i:=0 to Count-1 do begin
    if Value=TOleEnumArray(Enums)[i].Value then begin
      result := TOleEnumArray(Enums)[i].Name;
      exit;
    end;
  end;
  result := 'Unknown';
end;

function GetOleEnumText(const Enums; Count:integer; Value:longword) : string;
begin
  result := Format('%s ($%0.8x)', [GetOleEnumName(Enums, Count, Value), Value]);
end;

{------ }

function MsoAlignCmdStr(AValue: LongWord): string;
const
  Data : array[0..5] of TOleEnumItem = (
    (Name:'msoAlignLefts'; Value:$00000000),
    (Name:'msoAlignCenters'; Value:$00000001),
    (Name:'msoAlignRights'; Value:$00000002),
    (Name:'msoAlignTops'; Value:$00000003),
    (Name:'msoAlignMiddles'; Value:$00000004),
    (Name:'msoAlignBottoms'; Value:$00000005)
  );
begin
  Result := GetOleEnumText(Data, 6, AValue);
end;

function MsoAnimAccumulateStr(AValue: LongWord): string;
const
  Data : array[1..2] of TOleEnumItem = (
    (Name:'msoAnimAccumulateNone'; Value:$00000001),
    (Name:'msoAnimAccumulateAlways'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 2, AValue);
end;

function MsoAnimAdditiveStr(AValue: LongWord): string;
const
  Data : array[1..2] of TOleEnumItem = (
    (Name:'msoAnimAdditiveAddBase'; Value:$00000001),
    (Name:'msoAnimAdditiveAddSum'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 2, AValue);
end;

function MsoAnimAfterEffectStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'msoAnimAfterEffectMixed'; Value:$FFFFFFFF),
    (Name:'msoAnimAfterEffectNone'; Value:$00000000),
    (Name:'msoAnimAfterEffectDim'; Value:$00000001),
    (Name:'msoAnimAfterEffectHide'; Value:$00000002),
    (Name:'msoAnimAfterEffectHideOnNextClick'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoAnimateByLevelStr(AValue: Longword): string;
const
  Data : array[0..$1A+1] of TOleEnumItem = (
    (Name:'msoAnimateLevelMixed'; Value:$FFFFFFFF),
    (Name:'msoAnimateLevelNone'; Value:$00000000),
    (Name:'msoAnimateTextByAllLevels'; Value:$00000001),
    (Name:'msoAnimateTextByFirstLevel'; Value:$00000002),
    (Name:'msoAnimateTextBySecondLevel'; Value:$00000003),
    (Name:'msoAnimateTextByThirdLevel'; Value:$00000004),
    (Name:'msoAnimateTextByFourthLevel'; Value:$00000005),
    (Name:'msoAnimateTextByFifthLevel'; Value:$00000006),
    (Name:'msoAnimateChartAllAtOnce'; Value:$00000007),
    (Name:'msoAnimateChartByCategory'; Value:$00000008),
    (Name:'msoAnimateChartByCategoryElements'; Value:$00000009),
    (Name:'msoAnimateChartBySeries'; Value:$0000000A),
    (Name:'msoAnimateChartBySeriesElements'; Value:$0000000B),
    (Name:'msoAnimateDiagramAllAtOnce'; Value:$0000000C),
    (Name:'msoAnimateDiagramDepthByNode'; Value:$0000000D),
    (Name:'msoAnimateDiagramDepthByBranch'; Value:$0000000E),
    (Name:'msoAnimateDiagramBreadthByNode'; Value:$0000000F),
    (Name:'msoAnimateDiagramBreadthByLevel'; Value:$00000010),
    (Name:'msoAnimateDiagramClockwise'; Value:$00000011),
    (Name:'msoAnimateDiagramClockwiseIn'; Value:$00000012),
    (Name:'msoAnimateDiagramClockwiseOut'; Value:$00000013),
    (Name:'msoAnimateDiagramCounterClockwise'; Value:$00000014),
    (Name:'msoAnimateDiagramCounterClockwiseIn'; Value:$00000015),
    (Name:'msoAnimateDiagramCounterClockwiseOut'; Value:$00000016),
    (Name:'msoAnimateDiagramInByRing'; Value:$00000017),
    (Name:'msoAnimateDiagramOutByRing'; Value:$00000018),
    (Name:'msoAnimateDiagramUp'; Value:$00000019),
    (Name:'msoAnimateDiagramDown'; Value:$0000001A)
  );
begin
  Result := GetOleEnumText(Data, $1A+2, AValue);
end;

function MsoAnimationTypeStr(AValue: LongWord): string;
const
  data : array[1..35] of TOleEnumItem = (
    (Name:'msoAnimationIdle'; Value:$00000001),
    (Name:'msoAnimationGreeting'; Value:$00000002),
    (Name:'msoAnimationGoodbye'; Value:$00000003),
    (Name:'msoAnimationBeginSpeaking'; Value:$00000004),
    (Name:'msoAnimationRestPose'; Value:$00000005),
    (Name:'msoAnimationCharacterSuccessMajor'; Value:$00000006),
    (Name:'msoAnimationGetAttentionMajor'; Value:$0000000B),
    (Name:'msoAnimationGetAttentionMinor'; Value:$0000000C),
    (Name:'msoAnimationSearching'; Value:$0000000D),
    (Name:'msoAnimationPrinting'; Value:$00000012),
    (Name:'msoAnimationGestureRight'; Value:$00000013),
    (Name:'msoAnimationWritingNotingSomething'; Value:$00000016),
    (Name:'msoAnimationWorkingAtSomething'; Value:$00000017),
    (Name:'msoAnimationThinking'; Value:$00000018),
    (Name:'msoAnimationSendingMail'; Value:$00000019),
    (Name:'msoAnimationListensToComputer'; Value:$0000001A),
    (Name:'msoAnimationDisappear'; Value:$0000001F),
    (Name:'msoAnimationAppear'; Value:$00000020),
    (Name:'msoAnimationGetArtsy'; Value:$00000064),
    (Name:'msoAnimationGetTechy'; Value:$00000065),
    (Name:'msoAnimationGetWizardy'; Value:$00000066),
    (Name:'msoAnimationCheckingSomething'; Value:$00000067),
    (Name:'msoAnimationLookDown'; Value:$00000068),
    (Name:'msoAnimationLookDownLeft'; Value:$00000069),
    (Name:'msoAnimationLookDownRight'; Value:$0000006A),
    (Name:'msoAnimationLookLeft'; Value:$0000006B),
    (Name:'msoAnimationLookRight'; Value:$0000006C),
    (Name:'msoAnimationLookUp'; Value:$0000006D),
    (Name:'msoAnimationLookUpLeft'; Value:$0000006E),
    (Name:'msoAnimationLookUpRight'; Value:$0000006F),
    (Name:'msoAnimationSaving'; Value:$00000070),
    (Name:'msoAnimationGestureDown'; Value:$00000071),
    (Name:'msoAnimationGestureLeft'; Value:$00000072),
    (Name:'msoAnimationGestureUp'; Value:$00000073),
    (Name:'msoAnimationEmptyTrash'; Value:$00000074)
  );
begin
  Result := GetOleEnumText(Data, 35, AValue);
end;

function MsoAnimDirectionStr(AValue: LongWord): string;
const
  Data : array[0..$2C] of TOleEnumItem = (
    (Name:'msoAnimDirectionNone'; Value:$00000000),
    (Name:'msoAnimDirectionUp'; Value:$00000001),
    (Name:'msoAnimDirectionRight'; Value:$00000002),
    (Name:'msoAnimDirectionDown'; Value:$00000003),
    (Name:'msoAnimDirectionLeft'; Value:$00000004),
    (Name:'msoAnimDirectionOrdinalMask'; Value:$00000005),
    (Name:'msoAnimDirectionUpLeft'; Value:$00000006),
    (Name:'msoAnimDirectionUpRight'; Value:$00000007),
    (Name:'msoAnimDirectionDownRight'; Value:$00000008),
    (Name:'msoAnimDirectionDownLeft'; Value:$00000009),
    (Name:'msoAnimDirectionTop'; Value:$0000000A),
    (Name:'msoAnimDirectionBottom'; Value:$0000000B),
    (Name:'msoAnimDirectionTopLeft'; Value:$0000000C),
    (Name:'msoAnimDirectionTopRight'; Value:$0000000D),
    (Name:'msoAnimDirectionBottomRight'; Value:$0000000E),
    (Name:'msoAnimDirectionBottomLeft'; Value:$0000000F),
    (Name:'msoAnimDirectionHorizontal'; Value:$00000010),
    (Name:'msoAnimDirectionVertical'; Value:$00000011),
    (Name:'msoAnimDirectionAcross'; Value:$00000012),
    (Name:'msoAnimDirectionIn'; Value:$00000013),
    (Name:'msoAnimDirectionOut'; Value:$00000014),
    (Name:'msoAnimDirectionClockwise'; Value:$00000015),
    (Name:'msoAnimDirectionCounterclockwise'; Value:$00000016),
    (Name:'msoAnimDirectionHorizontalIn'; Value:$00000017),
    (Name:'msoAnimDirectionHorizontalOut'; Value:$00000018),
    (Name:'msoAnimDirectionVerticalIn'; Value:$00000019),
    (Name:'msoAnimDirectionVerticalOut'; Value:$0000001A),
    (Name:'msoAnimDirectionSlightly'; Value:$0000001B),
    (Name:'msoAnimDirectionCenter'; Value:$0000001C),
    (Name:'msoAnimDirectionInSlightly'; Value:$0000001D),
    (Name:'msoAnimDirectionInCenter'; Value:$0000001E),
    (Name:'msoAnimDirectionInBottom'; Value:$0000001F),
    (Name:'msoAnimDirectionOutSlightly'; Value:$00000020),
    (Name:'msoAnimDirectionOutCenter'; Value:$00000021),
    (Name:'msoAnimDirectionOutBottom'; Value:$00000022),
    (Name:'msoAnimDirectionFontBold'; Value:$00000023),
    (Name:'msoAnimDirectionFontItalic'; Value:$00000024),
    (Name:'msoAnimDirectionFontUnderline'; Value:$00000025),
    (Name:'msoAnimDirectionFontStrikethrough'; Value:$00000026),
    (Name:'msoAnimDirectionFontShadow'; Value:$00000027),
    (Name:'msoAnimDirectionFontAllCaps'; Value:$00000028),
    (Name:'msoAnimDirectionInstant'; Value:$00000029),
    (Name:'msoAnimDirectionGradual'; Value:$0000002A),
    (Name:'msoAnimDirectionCycleClockwise'; Value:$0000002B),
    (Name:'msoAnimDirectionCycleCounterclockwise'; Value:$0000002C)
  );
begin
  Result := GetOleEnumText(Data, $2C+1, AValue);
end;

function MsoAnimEffectAfterStr(AValue: LongWord): string;
const
  Data : array[1..4] of TOleEnumItem = (
    (Name:'msoAnimEffectAfterFreeze'; Value:$00000001),
    (Name:'msoAnimEffectAfterRemove'; Value:$00000002),
    (Name:'msoAnimEffectAfterHold'; Value:$00000003),
    (Name:'msoAnimEffectAfterTransition'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoAnimEffectRestartStr(AValue: LongWord): string;
const
  Data : array[1..3] of TOleEnumItem = (
    (Name:'msoAnimEffectRestartAlways'; Value:$00000001),
    (Name:'msoAnimEffectRestartWhenOff'; Value:$00000002),
    (Name:'msoAnimEffectRestartNever'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoAnimEffectStr(AValue: LongWord): string;
const
  Data : array[0..$95] of TOleEnumItem = (
    (Name:'msoAnimEffectCustom'; Value:$00000000),
    (Name:'msoAnimEffectAppear'; Value:$00000001),
    (Name:'msoAnimEffectFly'; Value:$00000002),
    (Name:'msoAnimEffectBlinds'; Value:$00000003),
    (Name:'msoAnimEffectBox'; Value:$00000004),
    (Name:'msoAnimEffectCheckerboard'; Value:$00000005),
    (Name:'msoAnimEffectCircle'; Value:$00000006),
    (Name:'msoAnimEffectCrawl'; Value:$00000007),
    (Name:'msoAnimEffectDiamond'; Value:$00000008),
    (Name:'msoAnimEffectDissolve'; Value:$00000009),
    (Name:'msoAnimEffectFade'; Value:$0000000A),
    (Name:'msoAnimEffectFlashOnce'; Value:$0000000B),
    (Name:'msoAnimEffectPeek'; Value:$0000000C),
    (Name:'msoAnimEffectPlus'; Value:$0000000D),
    (Name:'msoAnimEffectRandomBars'; Value:$0000000E),
    (Name:'msoAnimEffectSpiral'; Value:$0000000F),
    (Name:'msoAnimEffectSplit'; Value:$00000010),
    (Name:'msoAnimEffectStretch'; Value:$00000011),
    (Name:'msoAnimEffectStrips'; Value:$00000012),
    (Name:'msoAnimEffectSwivel'; Value:$00000013),
    (Name:'msoAnimEffectWedge'; Value:$00000014),
    (Name:'msoAnimEffectWheel'; Value:$00000015),
    (Name:'msoAnimEffectWipe'; Value:$00000016),
    (Name:'msoAnimEffectZoom'; Value:$00000017),
    (Name:'msoAnimEffectRandomEffects'; Value:$00000018),
    (Name:'msoAnimEffectBoomerang'; Value:$00000019),
    (Name:'msoAnimEffectBounce'; Value:$0000001A),
    (Name:'msoAnimEffectColorReveal'; Value:$0000001B),
    (Name:'msoAnimEffectCredits'; Value:$0000001C),
    (Name:'msoAnimEffectEaseIn'; Value:$0000001D),
    (Name:'msoAnimEffectFloat'; Value:$0000001E),
    (Name:'msoAnimEffectGrowAndTurn'; Value:$0000001F),
    (Name:'msoAnimEffectLightSpeed'; Value:$00000020),
    (Name:'msoAnimEffectPinwheel'; Value:$00000021),
    (Name:'msoAnimEffectRiseUp'; Value:$00000022),
    (Name:'msoAnimEffectSwish'; Value:$00000023),
    (Name:'msoAnimEffectThinLine'; Value:$00000024),
    (Name:'msoAnimEffectUnfold'; Value:$00000025),
    (Name:'msoAnimEffectWhip'; Value:$00000026),
    (Name:'msoAnimEffectAscend'; Value:$00000027),
    (Name:'msoAnimEffectCenterRevolve'; Value:$00000028),
    (Name:'msoAnimEffectFadedSwivel'; Value:$00000029),
    (Name:'msoAnimEffectDescend'; Value:$0000002A),
    (Name:'msoAnimEffectSling'; Value:$0000002B),
    (Name:'msoAnimEffectSpinner'; Value:$0000002C),
    (Name:'msoAnimEffectStretchy'; Value:$0000002D),
    (Name:'msoAnimEffectZip'; Value:$0000002E),
    (Name:'msoAnimEffectArcUp'; Value:$0000002F),
    (Name:'msoAnimEffectFadedZoom'; Value:$00000030),
    (Name:'msoAnimEffectGlide'; Value:$00000031),
    (Name:'msoAnimEffectExpand'; Value:$00000032),
    (Name:'msoAnimEffectFlip'; Value:$00000033),
    (Name:'msoAnimEffectShimmer'; Value:$00000034),
    (Name:'msoAnimEffectFold'; Value:$00000035),
    (Name:'msoAnimEffectChangeFillColor'; Value:$00000036),
    (Name:'msoAnimEffectChangeFont'; Value:$00000037),
    (Name:'msoAnimEffectChangeFontColor'; Value:$00000038),
    (Name:'msoAnimEffectChangeFontSize'; Value:$00000039),
    (Name:'msoAnimEffectChangeFontStyle'; Value:$0000003A),
    (Name:'msoAnimEffectGrowShrink'; Value:$0000003B),
    (Name:'msoAnimEffectChangeLineColor'; Value:$0000003C),
    (Name:'msoAnimEffectSpin'; Value:$0000003D),
    (Name:'msoAnimEffectTransparency'; Value:$0000003E),
    (Name:'msoAnimEffectBoldFlash'; Value:$0000003F),
    (Name:'msoAnimEffectBlast'; Value:$00000040),
    (Name:'msoAnimEffectBoldReveal'; Value:$00000041),
    (Name:'msoAnimEffectBrushOnColor'; Value:$00000042),
    (Name:'msoAnimEffectBrushOnUnderline'; Value:$00000043),
    (Name:'msoAnimEffectColorBlend'; Value:$00000044),
    (Name:'msoAnimEffectColorWave'; Value:$00000045),
    (Name:'msoAnimEffectComplementaryColor'; Value:$00000046),
    (Name:'msoAnimEffectComplementaryColor2'; Value:$00000047),
    (Name:'msoAnimEffectContrastingColor'; Value:$00000048),
    (Name:'msoAnimEffectDarken'; Value:$00000049),
    (Name:'msoAnimEffectDesaturate'; Value:$0000004A),
    (Name:'msoAnimEffectFlashBulb'; Value:$0000004B),
    (Name:'msoAnimEffectFlicker'; Value:$0000004C),
    (Name:'msoAnimEffectGrowWithColor'; Value:$0000004D),
    (Name:'msoAnimEffectLighten'; Value:$0000004E),
    (Name:'msoAnimEffectStyleEmphasis'; Value:$0000004F),
    (Name:'msoAnimEffectTeeter'; Value:$00000050),
    (Name:'msoAnimEffectVerticalGrow'; Value:$00000051),
    (Name:'msoAnimEffectWave'; Value:$00000052),
    (Name:'msoAnimEffectMediaPlay'; Value:$00000053),
    (Name:'msoAnimEffectMediaPause'; Value:$00000054),
    (Name:'msoAnimEffectMediaStop'; Value:$00000055),
    (Name:'msoAnimEffectPathCircle'; Value:$00000056),
    (Name:'msoAnimEffectPathRightTriangle'; Value:$00000057),
    (Name:'msoAnimEffectPathDiamond'; Value:$00000058),
    (Name:'msoAnimEffectPathHexagon'; Value:$00000059),
    (Name:'msoAnimEffectPath5PointStar'; Value:$0000005A),
    (Name:'msoAnimEffectPathCrescentMoon'; Value:$0000005B),
    (Name:'msoAnimEffectPathSquare'; Value:$0000005C),
    (Name:'msoAnimEffectPathTrapezoid'; Value:$0000005D),
    (Name:'msoAnimEffectPathHeart'; Value:$0000005E),
    (Name:'msoAnimEffectPathOctagon'; Value:$0000005F),
    (Name:'msoAnimEffectPath6PointStar'; Value:$00000060),
    (Name:'msoAnimEffectPathFootball'; Value:$00000061),
    (Name:'msoAnimEffectPathEqualTriangle'; Value:$00000062),
    (Name:'msoAnimEffectPathParallelogram'; Value:$00000063),
    (Name:'msoAnimEffectPathPentagon'; Value:$00000064),
    (Name:'msoAnimEffectPath4PointStar'; Value:$00000065),
    (Name:'msoAnimEffectPath8PointStar'; Value:$00000066),
    (Name:'msoAnimEffectPathTeardrop'; Value:$00000067),
    (Name:'msoAnimEffectPathPointyStar'; Value:$00000068),
    (Name:'msoAnimEffectPathCurvedSquare'; Value:$00000069),
    (Name:'msoAnimEffectPathCurvedX'; Value:$0000006A),
    (Name:'msoAnimEffectPathVerticalFigure8'; Value:$0000006B),
    (Name:'msoAnimEffectPathCurvyStar'; Value:$0000006C),
    (Name:'msoAnimEffectPathLoopdeLoop'; Value:$0000006D),
    (Name:'msoAnimEffectPathBuzzsaw'; Value:$0000006E),
    (Name:'msoAnimEffectPathHorizontalFigure8'; Value:$0000006F),
    (Name:'msoAnimEffectPathPeanut'; Value:$00000070),
    (Name:'msoAnimEffectPathFigure8Four'; Value:$00000071),
    (Name:'msoAnimEffectPathNeutron'; Value:$00000072),
    (Name:'msoAnimEffectPathSwoosh'; Value:$00000073),
    (Name:'msoAnimEffectPathBean'; Value:$00000074),
    (Name:'msoAnimEffectPathPlus'; Value:$00000075),
    (Name:'msoAnimEffectPathInvertedTriangle'; Value:$00000076),
    (Name:'msoAnimEffectPathInvertedSquare'; Value:$00000077),
    (Name:'msoAnimEffectPathLeft'; Value:$00000078),
    (Name:'msoAnimEffectPathTurnRight'; Value:$00000079),
    (Name:'msoAnimEffectPathArcDown'; Value:$0000007A),
    (Name:'msoAnimEffectPathZigzag'; Value:$0000007B),
    (Name:'msoAnimEffectPathSCurve2'; Value:$0000007C),
    (Name:'msoAnimEffectPathSineWave'; Value:$0000007D),
    (Name:'msoAnimEffectPathBounceLeft'; Value:$0000007E),
    (Name:'msoAnimEffectPathDown'; Value:$0000007F),
    (Name:'msoAnimEffectPathTurnUp'; Value:$00000080),
    (Name:'msoAnimEffectPathArcUp'; Value:$00000081),
    (Name:'msoAnimEffectPathHeartbeat'; Value:$00000082),
    (Name:'msoAnimEffectPathSpiralRight'; Value:$00000083),
    (Name:'msoAnimEffectPathWave'; Value:$00000084),
    (Name:'msoAnimEffectPathCurvyLeft'; Value:$00000085),
    (Name:'msoAnimEffectPathDiagonalDownRight'; Value:$00000086),
    (Name:'msoAnimEffectPathTurnDown'; Value:$00000087),
    (Name:'msoAnimEffectPathArcLeft'; Value:$00000088),
    (Name:'msoAnimEffectPathFunnel'; Value:$00000089),
    (Name:'msoAnimEffectPathSpring'; Value:$0000008A),
    (Name:'msoAnimEffectPathBounceRight'; Value:$0000008B),
    (Name:'msoAnimEffectPathSpiralLeft'; Value:$0000008C),
    (Name:'msoAnimEffectPathDiagonalUpRight'; Value:$0000008D),
    (Name:'msoAnimEffectPathTurnUpRight'; Value:$0000008E),
    (Name:'msoAnimEffectPathArcRight'; Value:$0000008F),
    (Name:'msoAnimEffectPathSCurve1'; Value:$00000090),
    (Name:'msoAnimEffectPathDecayingWave'; Value:$00000091),
    (Name:'msoAnimEffectPathCurvyRight'; Value:$00000092),
    (Name:'msoAnimEffectPathStairsDown'; Value:$00000093),
    (Name:'msoAnimEffectPathUp'; Value:$00000094),
    (Name:'msoAnimEffectPathRight'; Value:$00000095)
  );
begin
  Result := GetOleEnumText(Data, $95+1, AValue);
end;

function MsoAnimPropertyStr(AValue: LongWord): string;
const
  Data : array[1..43] of TOleEnumItem = (
    (Name:'msoAnimNone'; Value:$00000000),
    (Name:'msoAnimX'; Value:$00000001),
    (Name:'msoAnimY'; Value:$00000002),
    (Name:'msoAnimWidth'; Value:$00000003),
    (Name:'msoAnimHeight'; Value:$00000004),
    (Name:'msoAnimOpacity'; Value:$00000005),
    (Name:'msoAnimRotation'; Value:$00000006),
    (Name:'msoAnimColor'; Value:$00000007),
    (Name:'msoAnimVisibility'; Value:$00000008),
    (Name:'msoAnimTextFontBold'; Value:$00000064),
    (Name:'msoAnimTextFontColor'; Value:$00000065),
    (Name:'msoAnimTextFontEmboss'; Value:$00000066),
    (Name:'msoAnimTextFontItalic'; Value:$00000067),
    (Name:'msoAnimTextFontName'; Value:$00000068),
    (Name:'msoAnimTextFontShadow'; Value:$00000069),
    (Name:'msoAnimTextFontSize'; Value:$0000006A),
    (Name:'msoAnimTextFontSubscript'; Value:$0000006B),
    (Name:'msoAnimTextFontSuperscript'; Value:$0000006C),
    (Name:'msoAnimTextFontUnderline'; Value:$0000006D),
    (Name:'msoAnimTextFontStrikeThrough'; Value:$0000006E),
    (Name:'msoAnimTextBulletCharacter'; Value:$0000006F),
    (Name:'msoAnimTextBulletFontName'; Value:$00000070),
    (Name:'msoAnimTextBulletNumber'; Value:$00000071),
    (Name:'msoAnimTextBulletColor'; Value:$00000072),
    (Name:'msoAnimTextBulletRelativeSize'; Value:$00000073),
    (Name:'msoAnimTextBulletStyle'; Value:$00000074),
    (Name:'msoAnimTextBulletType'; Value:$00000075),
    (Name:'msoAnimShapePictureContrast'; Value:$000003E8),
    (Name:'msoAnimShapePictureBrightness'; Value:$000003E9),
    (Name:'msoAnimShapePictureGamma'; Value:$000003EA),
    (Name:'msoAnimShapePictureGrayscale'; Value:$000003EB),
    (Name:'msoAnimShapeFillOn'; Value:$000003EC),
    (Name:'msoAnimShapeFillColor'; Value:$000003ED),
    (Name:'msoAnimShapeFillOpacity'; Value:$000003EE),
    (Name:'msoAnimShapeFillBackColor'; Value:$000003EF),
    (Name:'msoAnimShapeLineOn'; Value:$000003F0),
    (Name:'msoAnimShapeLineColor'; Value:$000003F1),
    (Name:'msoAnimShapeShadowOn'; Value:$000003F2),
    (Name:'msoAnimShapeShadowType'; Value:$000003F3),
    (Name:'msoAnimShapeShadowColor'; Value:$000003F4),
    (Name:'msoAnimShapeShadowOpacity'; Value:$000003F5),
    (Name:'msoAnimShapeShadowOffsetX'; Value:$000003F6),
    (Name:'msoAnimShapeShadowOffsetY'; Value:$000003F7)
  );
begin
  Result := GetOleEnumText(Data, 43, AValue);
end;

function MsoAnimTextUnitEffectStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoAnimTextUnitEffectMixed'; Value:$FFFFFFFF),
    (Name:'msoAnimTextUnitEffectByParagraph'; Value:$00000000),
    (Name:'msoAnimTextUnitEffectByCharacter'; Value:$00000001),
    (Name:'msoAnimTextUnitEffectByWord'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoAnimTriggerTypeStr(AValue: LongWord): string;
const
  Data : array[0..5] of TOleEnumItem = (
    (Name:'msoAnimTriggerMixed'; Value:$FFFFFFFF),
    (Name:'msoAnimTriggerNone'; Value:$00000000),
    (Name:'msoAnimTriggerOnPageClick'; Value:$00000001),
    (Name:'msoAnimTriggerWithPrevious'; Value:$00000002),
    (Name:'msoAnimTriggerAfterPrevious'; Value:$00000003),
    (Name:'msoAnimTriggerOnShapeClick'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 6, AValue);
end;

function MsoAnimTypeStr(AValue: LongWord): string;
const
  Data : array[0..9] of TOleEnumItem = (
    (Name:'msoAnimTypeMixed'; Value:$FFFFFFFE),
    (Name:'msoAnimTypeNone'; Value:$00000000),
    (Name:'msoAnimTypeMotion'; Value:$00000001),
    (Name:'msoAnimTypeColor'; Value:$00000002),
    (Name:'msoAnimTypeScale'; Value:$00000003),
    (Name:'msoAnimTypeRotation'; Value:$00000004),
    (Name:'msoAnimTypeProperty'; Value:$00000005),
    (Name:'msoAnimTypeCommand'; Value:$00000006),
    (Name:'msoAnimTypeFilter'; Value:$00000007),
    (Name:'msoAnimTypeSet'; Value:$00000008)
  );
begin
  Result := GetOleEnumText(Data, 10, AValue);
end;

function MsoAppLanguageIDStr(AValue: LongWord): string;
const
  Data : array[1..5] of TOleEnumItem = (
    (Name:'msoLanguageIDInstall'; Value:$00000001),
    (Name:'msoLanguageIDUI'; Value:$00000002),
    (Name:'msoLanguageIDHelp'; Value:$00000003),
    (Name:'msoLanguageIDExeMode'; Value:$00000004),
    (Name:'msoLanguageIDUIPrevious'; Value:$00000005)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoArrowHeadLengthStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoArrowheadLengthMixed'; Value:$FFFFFFFE),
    (Name:'msoArrowheadShort'; Value:$00000001),
    (Name:'msoArrowheadLengthMedium'; Value:$00000002),
    (Name:'msoArrowheadLong'; Value:$00000003)
  );
begin
  result := GetOleEnumText(Data, 4, AValue);
end;

function MsoArrowHeadStyleStr(AValue: LongWord): string;
const
  Data : array[0..6] of TOleEnumItem = (
    (Name:'msoArrowheadStyleMixed'; Value:$FFFFFFFE),
    (Name:'msoArrowheadNone'; Value:$00000001),
    (Name:'msoArrowheadTriangle'; Value:$00000002),
    (Name:'msoArrowheadOpen'; Value:$00000003),
    (Name:'msoArrowheadStealth'; Value:$00000004),
    (Name:'msoArrowheadDiamond'; Value:$00000005),
    (Name:'msoArrowheadOval'; value:$00000006)
  );
begin
  result := GetOleEnumText(Data, 7, AValue);
end;

function MsoArrowHeadWidthStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoArrowheadWidthMixed'; Value:$FFFFFFFE),
    (Name:'msoArrowheadNarrow'; Value:$00000001),
    (Name:'msoArrowheadWidthMedium'; Value:$00000002),
    (Name:'msoArrowheadWide'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoAutoShapeTypeStr(AValue: LongWord): string;
const
  data : array[0..$8A] of TOleEnumItem = (
    (Name:'msoShapeMixed'; value:$FFFFFFFE),
    (Name:'msoShapeRectangle'; value:$00000001),
    (Name:'msoShapeParallelogram'; value:$00000002),
    (Name:'msoShapeTrapezoid'; value:$00000003),
    (Name:'msoShapeDiamond'; value:$00000004),
    (Name:'msoShapeRoundedRectangle'; value:$00000005),
    (Name:'msoShapeOctagon'; value:$00000006),
    (Name:'msoShapeIsoscelesTriangle'; value:$00000007),
    (Name:'msoShapeRightTriangle'; value:$00000008),
    (Name:'msoShapeOval'; value:$00000009),
    (Name:'msoShapeHexagon'; value:$0000000A),
    (Name:'msoShapeCross'; value:$0000000B),
    (Name:'msoShapeRegularPentagon'; value:$0000000C),
    (Name:'msoShapeCan'; value:$0000000D),
    (Name:'msoShapeCube'; value:$0000000E),
    (Name:'msoShapeBevel'; value:$0000000F),
    (Name:'msoShapeFoldedCorner'; value:$00000010),
    (Name:'msoShapeSmileyFace'; value:$00000011),
    (Name:'msoShapeDonut'; value:$00000012),
    (Name:'msoShapeNoSymbol'; value:$00000013),
    (Name:'msoShapeBlockArc'; value:$00000014),
    (Name:'msoShapeHeart'; value:$00000015),
    (Name:'msoShapeLightningBolt'; value:$00000016),
    (Name:'msoShapeSun'; value:$00000017),
    (Name:'msoShapeMoon'; value:$00000018),
    (Name:'msoShapeArc'; value:$00000019),
    (Name:'msoShapeDoubleBracket'; value:$0000001A),
    (Name:'msoShapeDoubleBrace'; value:$0000001B),
    (Name:'msoShapePlaque'; value:$0000001C),
    (Name:'msoShapeLeftBracket'; value:$0000001D),
    (Name:'msoShapeRightBracket'; value:$0000001E),
    (Name:'msoShapeLeftBrace'; value:$0000001F),
    (Name:'msoShapeRightBrace'; value:$00000020),
    (Name:'msoShapeRightArrow'; value:$00000021),
    (Name:'msoShapeLeftArrow'; value:$00000022),
    (Name:'msoShapeUpArrow'; value:$00000023),
    (Name:'msoShapeDownArrow'; value:$00000024),
    (Name:'msoShapeLeftRightArrow'; value:$00000025),
    (Name:'msoShapeUpDownArrow'; value:$00000026),
    (Name:'msoShapeQuadArrow'; value:$00000027),
    (Name:'msoShapeLeftRightUpArrow'; value:$00000028),
    (Name:'msoShapeBentArrow'; value:$00000029),
    (Name:'msoShapeUTurnArrow'; value:$0000002A),
    (Name:'msoShapeLeftUpArrow'; value:$0000002B),
    (Name:'msoShapeBentUpArrow'; value:$0000002C),
    (Name:'msoShapeCurvedRightArrow'; value:$0000002D),
    (Name:'msoShapeCurvedLeftArrow'; value:$0000002E),
    (Name:'msoShapeCurvedUpArrow'; value:$0000002F),
    (Name:'msoShapeCurvedDownArrow'; value:$00000030),
    (Name:'msoShapeStripedRightArrow'; value:$00000031),
    (Name:'msoShapeNotchedRightArrow'; value:$00000032),
    (Name:'msoShapePentagon'; value:$00000033),
    (Name:'msoShapeChevron'; value:$00000034),
    (Name:'msoShapeRightArrowCallout'; value:$00000035),
    (Name:'msoShapeLeftArrowCallout'; value:$00000036),
    (Name:'msoShapeUpArrowCallout'; value:$00000037),
    (Name:'msoShapeDownArrowCallout'; value:$00000038),
    (Name:'msoShapeLeftRightArrowCallout'; value:$00000039),
    (Name:'msoShapeUpDownArrowCallout'; value:$0000003A),
    (Name:'msoShapeQuadArrowCallout'; value:$0000003B),
    (Name:'msoShapeCircularArrow'; value:$0000003C),
    (Name:'msoShapeFlowchartProcess'; value:$0000003D),
    (Name:'msoShapeFlowchartAlternateProcess'; value:$0000003E),
    (Name:'msoShapeFlowchartDecision'; value:$0000003F),
    (Name:'msoShapeFlowchartData'; value:$00000040),
    (Name:'msoShapeFlowchartPredefinedProcess'; value:$00000041),
    (Name:'msoShapeFlowchartInternalStorage'; value:$00000042),
    (Name:'msoShapeFlowchartDocument'; value:$00000043),
    (Name:'msoShapeFlowchartMultidocument'; value:$00000044),
    (Name:'msoShapeFlowchartTerminator'; value:$00000045),
    (Name:'msoShapeFlowchartPreparation'; value:$00000046),
    (Name:'msoShapeFlowchartManualInput'; value:$00000047),
    (Name:'msoShapeFlowchartManualOperation'; value:$00000048),
    (Name:'msoShapeFlowchartConnector'; value:$00000049),
    (Name:'msoShapeFlowchartOffpageConnector'; value:$0000004A),
    (Name:'msoShapeFlowchartCard'; value:$0000004B),
    (Name:'msoShapeFlowchartPunchedTape'; value:$0000004C),
    (Name:'msoShapeFlowchartSummingJunction'; value:$0000004D),
    (Name:'msoShapeFlowchartOr'; value:$0000004E),
    (Name:'msoShapeFlowchartCollate'; value:$0000004F),
    (Name:'msoShapeFlowchartSort'; value:$00000050),
    (Name:'msoShapeFlowchartExtract'; value:$00000051),
    (Name:'msoShapeFlowchartMerge'; value:$00000052),
    (Name:'msoShapeFlowchartStoredData'; value:$00000053),
    (Name:'msoShapeFlowchartDelay'; value:$00000054),
    (Name:'msoShapeFlowchartSequentialAccessStorage'; value:$00000055),
    (Name:'msoShapeFlowchartMagneticDisk'; value:$00000056),
    (Name:'msoShapeFlowchartDirectAccessStorage'; value:$00000057),
    (Name:'msoShapeFlowchartDisplay'; value:$00000058),
    (Name:'msoShapeExplosion1'; value:$00000059),
    (Name:'msoShapeExplosion2'; value:$0000005A),
    (Name:'msoShape4pointStar'; value:$0000005B),
    (Name:'msoShape5pointStar'; value:$0000005C),
    (Name:'msoShape8pointStar'; value:$0000005D),
    (Name:'msoShape16pointStar'; value:$0000005E),
    (Name:'msoShape24pointStar'; value:$0000005F),
    (Name:'msoShape32pointStar'; value:$00000060),
    (Name:'msoShapeUpRibbon'; value:$00000061),
    (Name:'msoShapeDownRibbon'; value:$00000062),
    (Name:'msoShapeCurvedUpRibbon'; value:$00000063),
    (Name:'msoShapeCurvedDownRibbon'; value:$00000064),
    (Name:'msoShapeVerticalScroll'; value:$00000065),
    (Name:'msoShapeHorizontalScroll'; value:$00000066),
    (Name:'msoShapeWave'; value:$00000067),
    (Name:'msoShapeDoubleWave'; value:$00000068),
    (Name:'msoShapeRectangularCallout'; value:$00000069),
    (Name:'msoShapeRoundedRectangularCallout'; value:$0000006A),
    (Name:'msoShapeOvalCallout'; value:$0000006B),
    (Name:'msoShapeCloudCallout'; value:$0000006C),
    (Name:'msoShapeLineCallout1'; value:$0000006D),
    (Name:'msoShapeLineCallout2'; value:$0000006E),
    (Name:'msoShapeLineCallout3'; value:$0000006F),
    (Name:'msoShapeLineCallout4'; value:$00000070),
    (Name:'msoShapeLineCallout1AccentBar'; value:$00000071),
    (Name:'msoShapeLineCallout2AccentBar'; value:$00000072),
    (Name:'msoShapeLineCallout3AccentBar'; value:$00000073),
    (Name:'msoShapeLineCallout4AccentBar'; value:$00000074),
    (Name:'msoShapeLineCallout1NoBorder'; value:$00000075),
    (Name:'msoShapeLineCallout2NoBorder'; value:$00000076),
    (Name:'msoShapeLineCallout3NoBorder'; value:$00000077),
    (Name:'msoShapeLineCallout4NoBorder'; value:$00000078),
    (Name:'msoShapeLineCallout1BorderandAccentBar'; value:$00000079),
    (Name:'msoShapeLineCallout2BorderandAccentBar'; value:$0000007A),
    (Name:'msoShapeLineCallout3BorderandAccentBar'; value:$0000007B),
    (Name:'msoShapeLineCallout4BorderandAccentBar'; value:$0000007C),
    (Name:'msoShapeActionButtonCustom'; value:$0000007D),
    (Name:'msoShapeActionButtonHome'; value:$0000007E),
    (Name:'msoShapeActionButtonHelp'; value:$0000007F),
    (Name:'msoShapeActionButtonInformation'; value:$00000080),
    (Name:'msoShapeActionButtonBackorPrevious'; value:$00000081),
    (Name:'msoShapeActionButtonForwardorNext'; value:$00000082),
    (Name:'msoShapeActionButtonBeginning'; value:$00000083),
    (Name:'msoShapeActionButtonEnd'; value:$00000084),
    (Name:'msoShapeActionButtonReturn'; value:$00000085),
    (Name:'msoShapeActionButtonDocument'; value:$00000086),
    (Name:'msoShapeActionButtonSound'; value:$00000087),
    (Name:'msoShapeActionButtonMovie'; value:$00000088),
    (Name:'msoShapeBalloon'; value:$00000089),
    (Name:'msoShapeNotPrimitive'; value:$0000008A)
  );
begin
  Result := GetOleEnumText(Data, $8A+1, AValue);
end;

function MsoBlackWhiteModeStr(AValue: LongWord): string;
const
  Data : array[0..$A] of TOleEnumItem = (
    (Name:'msoBlackWhiteMixed'; value:$FFFFFFFE),
    (Name:'msoBlackWhiteAutomatic'; value:$00000001),
    (Name:'msoBlackWhiteGrayScale'; value:$00000002),
    (Name:'msoBlackWhiteLightGrayScale'; value:$00000003),
    (Name:'msoBlackWhiteInverseGrayScale'; value:$00000004),
    (Name:'msoBlackWhiteGrayOutline'; value:$00000005),
    (Name:'msoBlackWhiteBlackTextAndLine'; value:$00000006),
    (Name:'msoBlackWhiteHighContrast'; value:$00000007),
    (Name:'msoBlackWhiteBlack'; value:$00000008),
    (Name:'msoBlackWhiteWhite'; value:$00000009),
    (Name:'msoBlackWhiteDontShow'; value:$0000000A)
  );
begin
  Result := GetOleEnumText(Data, $0A+1, AValue);
end;

function MsoCalloutAngleTypeStr(AValue: LongWord): string;
const
  Data : array[0..5] of TOleEnumItem = (
    (Name:'msoCalloutAngleMixed'; Value:$FFFFFFFE),
    (Name:'msoCalloutAngleAutomatic'; Value:$00000001),
    (Name:'msoCalloutAngle30'; Value:$00000002),
    (Name:'msoCalloutAngle45'; Value:$00000003),
    (Name:'msoCalloutAngle60'; Value:$00000004),
    (Name:'msoCalloutAngle90'; Value:$00000005)
  );
begin
  Result := GetOleEnumText(Data, 6, AValue);
end;

function MsoCalloutDropTypeStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'msoCalloutDropMixed'; Value:$FFFFFFFE),
    (Name:'msoCalloutDropCustom'; Value:$00000001),
    (Name:'msoCalloutDropTop'; Value:$00000002),
    (Name:'msoCalloutDropCenter'; Value:$00000003),
    (Name:'msoCalloutDropBottom'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoCalloutTypeStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'msoCalloutMixed'; Value:$FFFFFFFE),
    (Name:'msoCalloutOne'; Value:$00000001),
    (Name:'msoCalloutTwo'; Value:$00000002),
    (Name:'msoCalloutThree'; Value:$00000003),
    (Name:'msoCalloutFour'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoColorTypeStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'msoColorTypeMixed'; Value:$FFFFFFFE),
    (Name:'msoColorTypeRGB'; Value:$00000001),
    (Name:'msoColorTypeScheme'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoConnectorTypeStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoConnectorTypeMixed'; Value:$FFFFFFFE),
    (Name:'msoConnectorStraight'; Value:$00000001),
    (Name:'msoConnectorElbow'; Value:$00000002),
    (Name:'msoConnectorCurve'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoDiagramTypeStr(AValue: LongWord): string;
const
  Data : array[0..6] of TOleEnumItem = (
    (Name:'msoDiagramMixed'; Value:$FFFFFFFE),
    (Name:'msoDiagramOrgChart'; Value:$00000001),
    (Name:'msoDiagramCycle'; Value:$00000002),
    (Name:'msoDiagramRadial'; Value:$00000003),
    (Name:'msoDiagramPyramid'; Value:$00000004),
    (Name:'msoDiagramVenn'; Value:$00000005),
    (Name:'msoDiagramTarget'; Value:$00000006)
  );
begin
  Result := GetOleEnumText(Data, 6, AValue);
end;

function MsoDistributeCmdStr(AValue: LongWord): string;
const
  Data : array[0..1] of TOleEnumItem = (
    (Name:'msoDistributeHorizontally'; Value:$00000000),
    (Name:'msoDistributeVertically'; Value:$00000001)
  );
begin
  Result := GetOleEnumText(Data, 2, AValue);
end;

function MsoEditingTypeStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoEditingAuto'; Value:$00000000),
    (Name:'msoEditingCorner'; Value:$00000001),
    (Name:'msoEditingSmooth'; Value:$00000002),
    (Name:'msoEditingSymmetric'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoFillTypeStr(AValue: LongWord): string;
const
  Data : array[0..6] of TOleEnumItem = (
    (Name:'msoFillMixed'; Value:$FFFFFFFE),
    (Name:'msoFillSolid'; Value:$00000001),
    (Name:'msoFillPatterned'; Value:$00000002),
    (Name:'msoFillGradient'; Value:$00000003),
    (Name:'msoFillTextured'; Value:$00000004),
    (Name:'msoFillBackground'; Value:$00000005),
    (Name:'msoFillPicture'; Value:$00000006)
  );
begin
  Result := GetOleEnumText(Data, 7, AValue);
end;

function MsoExtrusionColorTypeStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'msoExtrusionColorTypeMixed'; Value:$FFFFFFFE),
    (Name:'msoExtrusionColorAutomatic'; Value:$00000001),
    (Name:'msoExtrusionColorCustom'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoGradientColorTypeStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoGradientColorMixed'; Value:$FFFFFFFE),
    (Name:'msoGradientOneColor'; Value:$00000001),
    (Name:'msoGradientTwoColors'; Value:$00000002),
    (Name:'msoGradientPresetColors'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoGradientStyleStr(AValue: LongWord): string;
const
  Data : array[0..7] of TOleEnumItem = (
    (Name:'msoGradientMixed'; Value:$FFFFFFFE),
    (Name:'msoGradientHorizontal'; Value:$00000001),
    (Name:'msoGradientVertical'; Value:$00000002),
    (Name:'msoGradientDiagonalUp'; Value:$00000003),
    (Name:'msoGradientDiagonalDown'; Value:$00000004),
    (Name:'msoGradientFromCorner'; Value:$00000005),
    (Name:'msoGradientFromTitle'; Value:$00000006),
    (Name:'msoGradientFromCenter'; Value:$00000007)
  );
begin
  Result := GetOleEnumText(Data, 8, AValue);
end;

function MsoHorizontalAnchorStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'msoHorizontalAnchorMixed'; value:$FFFFFFFE),
    (Name:'msoAnchorNone'; value:$00000001),
    (Name:'msoAnchorCenter'; value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoLineDashStyleStr(AValue: LongWord): string;
const
  Data : array[0..8] of TOleEnumItem = (
    (Name:'msoLineDashStyleMixed'; Value:$FFFFFFFE),
    (Name:'msoLineSolid'; Value:$00000001),
    (Name:'msoLineSquareDot'; Value:$00000002),
    (Name:'msoLineRoundDot'; Value:$00000003),
    (Name:'msoLineDash'; Value:$00000004),
    (Name:'msoLineDashDot'; Value:$00000005),
    (Name:'msoLineDashDotDot'; Value:$00000006),
    (Name:'msoLineLongDash'; Value:$00000007),
    (Name:'msoLineLongDashDot'; Value:$00000008)
  );
begin
  Result := GetOleEnumText(Data, 9, AValue);
end;

function MsoLineStyleStr(AValue: LongWord): string;
const
  Data : array[0..5] of TOleEnumItem = (
    (Name:'msoLineStyleMixed'; Value:$FFFFFFFE),
    (Name:'msoLineSingle'; Value:$00000001),
    (Name:'msoLineThinThin'; Value:$00000002),
    (Name:'msoLineThinThick'; Value:$00000003),
    (Name:'msoLineThickThin'; Value:$00000004),
    (Name:'msoLineThickBetweenThin'; Value:$00000005)
  );
begin
   Result := GetOleEnumText(Data, 6, AValue);
end;

function MsoOrientationStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'msoOrientationMixed'; Value:$FFFFFFFE),
    (Name:'msoOrientationHorizontal'; Value:$00000001),
    (Name:'msoOrientationVertical'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoPatternTypeStr(AValue: LongWord): string;
const
  Data : array[0..$30] of TOleEnumItem = (
    (Name:'msoPatternMixed'; value:$FFFFFFFE),
    (Name:'msoPattern5Percent'; value:$00000001),
    (Name:'msoPattern10Percent'; value:$00000002),
    (Name:'msoPattern20Percent'; value:$00000003),
    (Name:'msoPattern25Percent'; value:$00000004),
    (Name:'msoPattern30Percent'; value:$00000005),
    (Name:'msoPattern40Percent'; value:$00000006),
    (Name:'msoPattern50Percent'; value:$00000007),
    (Name:'msoPattern60Percent'; value:$00000008),
    (Name:'msoPattern70Percent'; value:$00000009),
    (Name:'msoPattern75Percent'; value:$0000000A),
    (Name:'msoPattern80Percent'; value:$0000000B),
    (Name:'msoPattern90Percent'; value:$0000000C),
    (Name:'msoPatternDarkHorizontal'; value:$0000000D),
    (Name:'msoPatternDarkVertical'; value:$0000000E),
    (Name:'msoPatternDarkDownwardDiagonal'; value:$0000000F),
    (Name:'msoPatternDarkUpwardDiagonal'; value:$00000010),
    (Name:'msoPatternSmallCheckerBoard'; value:$00000011),
    (Name:'msoPatternTrellis'; value:$00000012),
    (Name:'msoPatternLightHorizontal'; value:$00000013),
    (Name:'msoPatternLightVertical'; value:$00000014),
    (Name:'msoPatternLightDownwardDiagonal'; value:$00000015),
    (Name:'msoPatternLightUpwardDiagonal'; value:$00000016),
    (Name:'msoPatternSmallGrid'; value:$00000017),
    (Name:'msoPatternDottedDiamond'; value:$00000018),
    (Name:'msoPatternWideDownwardDiagonal'; value:$00000019),
    (Name:'msoPatternWideUpwardDiagonal'; value:$0000001A),
    (Name:'msoPatternDashedUpwardDiagonal'; value:$0000001B),
    (Name:'msoPatternDashedDownwardDiagonal'; value:$0000001C),
    (Name:'msoPatternNarrowVertical'; value:$0000001D),
    (Name:'msoPatternNarrowHorizontal'; value:$0000001E),
    (Name:'msoPatternDashedVertical'; value:$0000001F),
    (Name:'msoPatternDashedHorizontal'; value:$00000020),
    (Name:'msoPatternLargeConfetti'; value:$00000021),
    (Name:'msoPatternLargeGrid'; value:$00000022),
    (Name:'msoPatternHorizontalBrick'; value:$00000023),
    (Name:'msoPatternLargeCheckerBoard'; value:$00000024),
    (Name:'msoPatternSmallConfetti'; value:$00000025),
    (Name:'msoPatternZigZag'; value:$00000026),
    (Name:'msoPatternSolidDiamond'; value:$00000027),
    (Name:'msoPatternDiagonalBrick'; value:$00000028),
    (Name:'msoPatternOutlinedDiamond'; value:$00000029),
    (Name:'msoPatternPlaid'; value:$0000002A),
    (Name:'msoPatternSphere'; value:$0000002B),
    (Name:'msoPatternWeave'; value:$0000002C),
    (Name:'msoPatternDottedGrid'; value:$0000002D),
    (Name:'msoPatternDivot'; value:$0000002E),
    (Name:'msoPatternShingle'; value:$0000002F),
    (Name:'msoPatternWave'; value:$00000030)
  );
begin
  Result := GetOleEnumText(Data, $30+1, AValue);
end;

function MsoPictureColorTypeStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'msoPictureMixed'; Value:$FFFFFFFE),
    (Name:'msoPictureAutomatic'; Value:$00000001),
    (Name:'msoPictureGrayscale'; Value:$00000002),
    (Name:'msoPictureBlackAndWhite'; Value:$00000003),
    (Name:'msoPictureWatermark'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoPresetExtrusionDirectionStr(AValue: LongWord): string;
const
  Data : array[0..9] of TOleEnumItem = (
    (Name:'msoPresetExtrusionDirectionMixed'; Value:$FFFFFFFE),
    (Name:'msoExtrusionBottomRight'; Value:$00000001),
    (Name:'msoExtrusionBottom'; Value:$00000002),
    (Name:'msoExtrusionBottomLeft'; Value:$00000003),
    (Name:'msoExtrusionRight'; Value:$00000004),
    (Name:'msoExtrusionNone'; Value:$00000005),
    (Name:'msoExtrusionLeft'; Value:$00000006),
    (Name:'msoExtrusionTopRight'; Value:$00000007),
    (Name:'msoExtrusionTop'; Value:$00000008),
    (Name:'msoExtrusionTopLeft'; Value:$00000009)
  );
begin
  Result := GetOleEnumText(Data, 10, AValue);
end;

function MsoPresetGradientTypeStr(AValue: LongWord): string;
const
  Data : array[0..$18] of TOleEnumItem = (
    (Name:'msoPresetGradientMixed'; Value:$FFFFFFFE),
    (Name:'msoGradientEarlySunset'; Value:$00000001),
    (Name:'msoGradientLateSunset'; Value:$00000002),
    (Name:'msoGradientNightfall'; Value:$00000003),
    (Name:'msoGradientDaybreak'; Value:$00000004),
    (Name:'msoGradientHorizon'; Value:$00000005),
    (Name:'msoGradientDesert'; Value:$00000006),
    (Name:'msoGradientOcean'; Value:$00000007),
    (Name:'msoGradientCalmWater'; Value:$00000008),
    (Name:'msoGradientFire'; Value:$00000009),
    (Name:'msoGradientFog'; Value:$0000000A),
    (Name:'msoGradientMoss'; Value:$0000000B),
    (Name:'msoGradientPeacock'; Value:$0000000C),
    (Name:'msoGradientWheat'; Value:$0000000D),
    (Name:'msoGradientParchment'; Value:$0000000E),
    (Name:'msoGradientMahogany'; Value:$0000000F),
    (Name:'msoGradientRainbow'; Value:$00000010),
    (Name:'msoGradientRainbowII'; Value:$00000011),
    (Name:'msoGradientGold'; Value:$00000012),
    (Name:'msoGradientGoldII'; Value:$00000013),
    (Name:'msoGradientBrass'; Value:$00000014),
    (Name:'msoGradientChrome'; Value:$00000015),
    (Name:'msoGradientChromeII'; Value:$00000016),
    (Name:'msoGradientSilver'; Value:$00000017),
    (Name:'msoGradientSapphire'; Value:$00000018)
  );
begin
  Result := GetOleEnumText(Data, $18+1, AValue);
end;

function MsoPresetLightingDirectionStr(AValue: LongWord): string;
const
  Data : array[0..9] of TOleEnumItem = (
    (Name:'msoPresetLightingDirectionMixed'; Value:$FFFFFFFE),
    (Name:'msoLightingTopLeft'; Value:$00000001),
    (Name:'msoLightingTop'; Value:$00000002),
    (Name:'msoLightingTopRight'; Value:$00000003),
    (Name:'msoLightingLeft'; Value:$00000004),
    (Name:'msoLightingNone'; Value:$00000005),
    (Name:'msoLightingRight'; Value:$00000006),
    (Name:'msoLightingBottomLeft'; Value:$00000007),
    (Name:'msoLightingBottom'; Value:$00000008),
    (Name:'msoLightingBottomRight'; Value:$00000009)
  );
begin
  Result := GetOleEnumText(Data, 10, AValue);
end;

function MsoPresetLightingSoftnessStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'msoPresetLightingSoftnessMixed'; Value:$FFFFFFFE),
    (Name:'msoLightingDim'; Value:$00000001),
    (Name:'msoLightingNormal'; Value:$00000002),
    (Name:'msoLightingBright'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function MsoPresetMaterialStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'msoPresetMaterialMixed'; Value:$FFFFFFFE),
    (Name:'msoMaterialMatte'; Value:$00000001),
    (Name:'msoMaterialPlastic'; Value:$00000002),
    (Name:'msoMaterialMetal'; Value:$00000003),
    (Name:'msoMaterialWireFrame'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoPresetTextureStr(AValue: LongWord): string;
const
  Data : array[0..$18] of TOleEnumItem = (
    (Name:'msoPresetTextureMixed'; Value:$FFFFFFFE),
    (Name:'msoTexturePapyrus'; Value:$00000001),
    (Name:'msoTextureCanvas'; Value:$00000002),
    (Name:'msoTextureDenim'; Value:$00000003),
    (Name:'msoTextureWovenMat'; Value:$00000004),
    (Name:'msoTextureWaterDroplets'; Value:$00000005),
    (Name:'msoTexturePaperBag'; Value:$00000006),
    (Name:'msoTextureFishFossil'; Value:$00000007),
    (Name:'msoTextureSand'; Value:$00000008),
    (Name:'msoTextureGreenMarble'; Value:$00000009),
    (Name:'msoTextureWhiteMarble'; Value:$0000000A),
    (Name:'msoTextureBrownMarble'; Value:$0000000B),
    (Name:'msoTextureGranite'; Value:$0000000C),
    (Name:'msoTextureNewsprint'; Value:$0000000D),
    (Name:'msoTextureRecycledPaper'; Value:$0000000E),
    (Name:'msoTextureParchment'; Value:$0000000F),
    (Name:'msoTextureStationery'; Value:$00000010),
    (Name:'msoTextureBlueTissuePaper'; Value:$00000011),
    (Name:'msoTexturePinkTissuePaper'; Value:$00000012),
    (Name:'msoTexturePurpleMesh'; Value:$00000013),
    (Name:'msoTextureBouquet'; Value:$00000014),
    (Name:'msoTextureCork'; Value:$00000015),
    (Name:'msoTextureWalnut'; Value:$00000016),
    (Name:'msoTextureOak'; Value:$00000017),
    (Name:'msoTextureMediumWood'; Value:$00000018)
  );
begin
  Result := GetOleEnumText(Data, $18+1, AValue);
end;

function MsoPresetThreeDFormatStr(AValue: LongWord): string;
const
  Data : array[0..$14] of TOleEnumItem = (
    (Name:'msoPresetThreeDFormatMixed'; Value:$FFFFFFFE),
    (Name:'msoThreeD1'; Value:$00000001),
    (Name:'msoThreeD2'; Value:$00000002),
    (Name:'msoThreeD3'; Value:$00000003),
    (Name:'msoThreeD4'; Value:$00000004),
    (Name:'msoThreeD5'; Value:$00000005),
    (Name:'msoThreeD6'; Value:$00000006),
    (Name:'msoThreeD7'; Value:$00000007),
    (Name:'msoThreeD8'; Value:$00000008),
    (Name:'msoThreeD9'; Value:$00000009),
    (Name:'msoThreeD10'; Value:$0000000A),
    (Name:'msoThreeD11'; Value:$0000000B),
    (Name:'msoThreeD12'; Value:$0000000C),
    (Name:'msoThreeD13'; Value:$0000000D),
    (Name:'msoThreeD14'; Value:$0000000E),
    (Name:'msoThreeD15'; Value:$0000000F),
    (Name:'msoThreeD16'; Value:$00000010),
    (Name:'msoThreeD17'; Value:$00000011),
    (Name:'msoThreeD18'; Value:$00000012),
    (Name:'msoThreeD19'; Value:$00000013),
    (Name:'msoThreeD20'; Value:$00000014)
  );
begin
  Result := GetOleEnumText(Data, $14+1, AValue);
end;

function MsoScaleFromStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'msoScaleFromTopLeft'; Value:$00000000),
    (Name:'msoScaleFromMiddle'; Value:$00000001),
    (Name:'msoScaleFromBottomRight'; Value:$00000002)
  );
begin
 Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoSegmentTypeStr(AValue: LongWord): string;
const
  Data : array[0..1] of TOleEnumItem = (
    (Name:'msoSegmentLine'; Value:$00000000),
    (Name:'msoSegmentCurve'; Value:$00000001)
  );
begin
  Result := GetOleEnumText(Data, 2, AValue);
end;

function MsoShadowTypeStr(AValue: LongWord): string;
const
  Data : array[0..$14] of TOleEnumItem = (
    (Name:'msoShadowMixed'; Value:$FFFFFFFE),
    (Name:'msoShadow1'; Value:$00000001),
    (Name:'msoShadow2'; Value:$00000002),
    (Name:'msoShadow3'; Value:$00000003),
    (Name:'msoShadow4'; Value:$00000004),
    (Name:'msoShadow5'; Value:$00000005),
    (Name:'msoShadow6'; Value:$00000006),
    (Name:'msoShadow7'; Value:$00000007),
    (Name:'msoShadow8'; Value:$00000008),
    (Name:'msoShadow9'; Value:$00000009),
    (Name:'msoShadow10'; Value:$0000000A),
    (Name:'msoShadow11'; Value:$0000000B),
    (Name:'msoShadow12'; Value:$0000000C),
    (Name:'msoShadow13'; Value:$0000000D),
    (Name:'msoShadow14'; Value:$0000000E),
    (Name:'msoShadow15'; Value:$0000000F),
    (Name:'msoShadow16'; Value:$00000010),
    (Name:'msoShadow17'; Value:$00000011),
    (Name:'msoShadow18'; Value:$00000012),
    (Name:'msoShadow19'; Value:$00000013),
    (Name:'msoShadow20'; Value:$00000014)
  );
begin
  Result := GetOleEnumText(Data, $14+1, AValue);
end;

function MsoShapeTypeStr(AValue: Longword): string;
const
  Data : array[0..$13] of TOleEnumItem = (
    (Name:'msoShapeTypeMixed'; Value:$FFFFFFFE),
    (Name:'msoAutoShape'; Value:$00000001),
    (Name:'msoCallout'; Value:$00000002),
    (Name:'msoChart'; Value:$00000003),
    (Name:'msoComment'; Value:$00000004),
    (Name:'msoFreeform'; Value:$00000005),
    (Name:'msoGroup'; Value:$00000006),
    (Name:'msoEmbeddedOLEObject'; Value:$00000007),
    (Name:'msoFormControl'; Value:$00000008),
    (Name:'msoLine'; Value:$00000009),
    (Name:'msoLinkedOLEObject'; Value:$0000000A),
    (Name:'msoLinkedPicture'; Value:$0000000B),
    (Name:'msoOLEControlObject'; Value:$0000000C),
    (Name:'msoPicture'; Value:$0000000D),
    (Name:'msoPlaceholder'; Value:$0000000E),
    (Name:'msoTextEffect'; Value:$0000000F),
    (Name:'msoMedia'; Value:$00000010),
    (Name:'msoTextBox'; Value:$00000011),
    (Name:'msoScriptAnchor'; Value:$00000012),
    (Name:'msoTable'; Value:$00000013)
  );
begin
  Result := GetOleEnumText(Data, $13+1, AValue);
end;

function MsoTextEffectAlignmentStr(AValue: LongWord): string;
const
  Data : array[0..6] of TOleEnumItem = (
    (Name:'msoTextEffectAlignmentMixed'; Value:$FFFFFFFE),
    (Name:'msoTextEffectAlignmentLeft'; Value:$00000001),
    (Name:'msoTextEffectAlignmentCentered'; Value:$00000002),
    (Name:'msoTextEffectAlignmentRight'; Value:$00000003),
    (Name:'msoTextEffectAlignmentLetterJustify'; Value:$00000004),
    (Name:'msoTextEffectAlignmentWordJustify'; Value:$00000005),
    (Name:'msoTextEffectAlignmentStretchJustify'; Value:$00000006)
  );
begin
  Result := GetOleEnumText(Data, 7, AValue);
end;

function MsoTextEffectShapeStr(AValue: LongWord): string;
const
  data : array[0..$28] of TOleEnumItem = (
    (Name:'msoTextEffectShapeMixed'; Value:$FFFFFFFE),
    (Name:'msoTextEffectShapePlainText'; Value:$00000001),
    (Name:'msoTextEffectShapeStop'; Value:$00000002),
    (Name:'msoTextEffectShapeTriangleUp'; Value:$00000003),
    (Name:'msoTextEffectShapeTriangleDown'; Value:$00000004),
    (Name:'msoTextEffectShapeChevronUp'; Value:$00000005),
    (Name:'msoTextEffectShapeChevronDown'; Value:$00000006),
    (Name:'msoTextEffectShapeRingInside'; Value:$00000007),
    (Name:'msoTextEffectShapeRingOutside'; Value:$00000008),
    (Name:'msoTextEffectShapeArchUpCurve'; Value:$00000009),
    (Name:'msoTextEffectShapeArchDownCurve'; Value:$0000000A),
    (Name:'msoTextEffectShapeCircleCurve'; Value:$0000000B),
    (Name:'msoTextEffectShapeButtonCurve'; Value:$0000000C),
    (Name:'msoTextEffectShapeArchUpPour'; Value:$0000000D),
    (Name:'msoTextEffectShapeArchDownPour'; Value:$0000000E),
    (Name:'msoTextEffectShapeCirclePour'; Value:$0000000F),
    (Name:'msoTextEffectShapeButtonPour'; Value:$00000010),
    (Name:'msoTextEffectShapeCurveUp'; Value:$00000011),
    (Name:'msoTextEffectShapeCurveDown'; Value:$00000012),
    (Name:'msoTextEffectShapeCanUp'; Value:$00000013),
    (Name:'msoTextEffectShapeCanDown'; Value:$00000014),
    (Name:'msoTextEffectShapeWave1'; Value:$00000015),
    (Name:'msoTextEffectShapeWave2'; Value:$00000016),
    (Name:'msoTextEffectShapeDoubleWave1'; Value:$00000017),
    (Name:'msoTextEffectShapeDoubleWave2'; Value:$00000018),
    (Name:'msoTextEffectShapeInflate'; Value:$00000019),
    (Name:'msoTextEffectShapeDeflate'; Value:$0000001A),
    (Name:'msoTextEffectShapeInflateBottom'; Value:$0000001B),
    (Name:'msoTextEffectShapeDeflateBottom'; Value:$0000001C),
    (Name:'msoTextEffectShapeInflateTop'; Value:$0000001D),
    (Name:'msoTextEffectShapeDeflateTop'; Value:$0000001E),
    (Name:'msoTextEffectShapeDeflateInflate'; Value:$0000001F),
    (Name:'msoTextEffectShapeDeflateInflateDeflate'; Value:$00000020),
    (Name:'msoTextEffectShapeFadeRight'; Value:$00000021),
    (Name:'msoTextEffectShapeFadeLeft'; Value:$00000022),
    (Name:'msoTextEffectShapeFadeUp'; Value:$00000023),
    (Name:'msoTextEffectShapeFadeDown'; Value:$00000024),
    (Name:'msoTextEffectShapeSlantUp'; Value:$00000025),
    (Name:'msoTextEffectShapeSlantDown'; Value:$00000026),
    (Name:'msoTextEffectShapeCascadeUp'; Value:$00000027),
    (Name:'msoTextEffectShapeCascadeDown'; Value:$00000028)
  );
begin
  Result := GetOleEnumText(Data, $28+1, AValue);
end;

function MsoTextOrientationStr(AValue: LongWord): string;
const
  Data : array[0..6] of TOleEnumItem = (
    (Name:'msoTextOrientationMixed'; Value:$FFFFFFFE),
    (Name:'msoTextOrientationHorizontal'; Value:$00000001),
    (Name:'msoTextOrientationUpward'; Value:$00000002),
    (Name:'msoTextOrientationDownward'; Value:$00000003),
    (Name:'msoTextOrientationVerticalFarEast'; Value:$00000004),
    (Name:'msoTextOrientationVertical'; Value:$00000005),
    (Name:'msoTextOrientationHorizontalRotatedFarEast'; Value:$00000006)
  );
begin
  Result := GetOleEnumText(Data, 7, AValue);
end;

function MsoTextureTypeStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'msoTextureTypeMixed'; Value:$FFFFFFFE),
    (Name:'msoTexturePreset'; Value:$00000001),
    (Name:'msoTextureUserDefined'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function MsoTriStateStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'msoTrue'; Value:$FFFFFFFF),
    (Name:'msoFalse'; Value:$00000000),
    (Name:'msoCTrue'; Value:$0000000),
    (Name:'msoTriStateToggle'; Value:$FFFFFFFD),
    (Name:'msoTriStateMixed'; Value:$FFFFFFFE)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function MsoVerticalAnchorStr(AValue: LongWord): string;
const
  Data : array[0..5] of TOleEnumItem = (
    (Name:'msoVerticalAnchorMixed'; Value:$FFFFFFFE),
    (Name:'msoAnchorTop'; Value:$00000001),
    (Name:'msoAnchorTopBaseline'; Value:$00000002),
    (Name:'msoAnchorMiddle'; Value:$00000003),
    (Name:'msoAnchorBottom'; Value:$00000004),
    (Name:'msoAnchorBottomBaseLine'; Value:$00000005)
  );
begin
  Result := GetOleEnumText(Data, 6, AValue);
end;

function MsoZOrderCmdStr(AValue: LongWord): string;
const
  Data : array[0..5] of TOleEnumItem = (
    (Name:'msoBringToFront'; Value:$00000000),
    (Name:'msoSendToBack'; Value:$00000001),
    (Name:'msoBringForward'; Value:$00000002),
    (Name:'msoSendBackward'; Value:$00000003),
    (Name:'msoBringInFrontOfText'; Value:$00000004),
    (Name:'msoSendBehindText'; Value:$00000005)
  );
begin
  Result := GetOleEnumText(Data, 6, AValue);
end;

{ -----------------------------------------------------------------------------}

function PpActionTypeStr(AValue: LongWord): string;
const
  Data : array[0..$0C+1] of TOleEnumItem = (
    (Name:'PpActionMixed'; Value:$FFFFFFFE),
    (Name:'PpActionNone'; Value:$00000000),
    (Name:'PpActionNextSlide'; Value:$00000001),
    (Name:'PpActionPreviousSlide'; Value:$00000002),
    (Name:'PpActionFirstSlide'; Value:$00000003),
    (Name:'PpActionLastSlide'; Value:$00000004),
    (Name:'PpActionLastSlideViewed'; Value:$00000005),
    (Name:'PpActionEndShow'; Value:$00000006),
    (Name:'PpActionHyperlink'; Value:$00000007),
    (Name:'PpActionRunMacro'; Value:$00000008),
    (Name:'PpActionRunProgram'; Value:$00000009),
    (Name:'PpActionNamedSlideShow'; Value:$0000000A),
    (Name:'PpActionOLEVerb'; Value:$0000000B),
    (Name:'PpActionPlay'; Value:$0000000C)
  );
begin
  Result := GetOleEnumText(Data, $C+2, AValue);
end;

function PpAdvanceModeStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'PpAdvanceModeMixed'; Value:$FFFFFFFE),
    (Name:'PpAdvanceOnClick'; Value:$00000001),
    (Name:'PpAdvanceOnTime'; Value:$00000002)
  );
begin
  result := GetOleEnumText(Data, 3, AValue);
end;

function PpAfterEffectStr(AValue: Longword): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'PpAfterEffectMixed'; Value:$FFFFFFFE),
    (Name:'PpAfterEffectNothing'; Value:$00000000),
    (Name:'PpAfterEffectHide'; Value:$00000001),
    (Name:'PpAfterEffectDim'; Value:$00000002),
    (Name:'PpAfterEffectHideOnClick'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function PpAutoSizeStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'PpAutoSizeMixed'; Value:$FFFFFFFE),
    (Name:'PpAutoSizeNone'; Value:$00000000),
    (Name:'PpAutoSizeShapeToFitText'; Value:$00000001)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function PpChartUnitEffectStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'PpAnimateChartMixed'; Value:$FFFFFFFE),
    (Name:'PpAnimateBySeries'; Value:$00000001),
    (Name:'PpAnimateByCategory'; Value:$00000002),
    (Name:'PpAnimateBySeriesElements'; Value:$00000003),
    (Name:'PpAnimateByCategoryElements'; Value:$00000004)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function PpColorSchemeIndexStr(AValue: LongWord): string;
const
  Data : array[0..9] of TOleEnumItem = (
    (Name:'PpSchemeColorMixed'; Value:$FFFFFFFE),
    (Name:'PpNotSchemeColor'; Value:$00000000),
    (Name:'PpBackground'; Value:$00000001),
    (Name:'PpForeground'; Value:$00000002),
    (Name:'PpShadow'; Value:$00000003),
    (Name:'PpTitle'; Value:$00000004),
    (Name:'PpFill'; Value:$00000005),
    (Name:'PpAccent1'; Value:$00000006),
    (Name:'PpAccent2'; Value:$00000007),
    (Name:'PpAccent3'; Value:$00000008)
  );
begin
  Result := GetOleEnumText(Data, 10, AValue);
end;

function PpDateTimeFormatStr(AValue: LongWord): string;
const
  Data : array[0..$0D] of TOleEnumItem = (
    (Name:'PpDateTimeFormatMixed'; Value:$FFFFFFFE),
    (Name:'PpDateTimeMdyy'; Value:$00000001),
    (Name:'PpDateTimeddddMMMMddyyyy'; Value:$00000002),
    (Name:'PpDateTimedMMMMyyyy'; Value:$00000003),
    (Name:'PpDateTimeMMMMdyyyy'; Value:$00000004),
    (Name:'PpDateTimedMMMyy'; Value:$00000005),
    (Name:'PpDateTimeMMMMyy'; Value:$00000006),
    (Name:'PpDateTimeMMyy'; Value:$00000007),
    (Name:'PpDateTimeMMddyyHmm'; Value:$00000008),
    (Name:'PpDateTimeMMddyyhmmAMPM'; Value:$00000009),
    (Name:'PpDateTimeHmm'; Value:$0000000A),
    (Name:'PpDateTimeHmmss'; Value:$0000000B),
    (Name:'PpDateTimehmmAMPM'; Value:$0000000C),
    (Name:'PpDateTimehmmssAMPM'; Value:$0000000D)
  );
begin
  Result := GetOleEnumText(Data, $D+1, AValue);
end;

function PpEntryEffectStr(AValue: LongWord): string;
const
  Data : array[0..79] of TOleEnumItem = (
    (Name:'PpEffectMixed'; value:$FFFFFFFE),
    (Name:'PpEffectNone'; value:$00000000),
    (Name:'PpEffectCut'; value:$00000101),
    (Name:'PpEffectCutThroughBlack'; value:$00000102),
    (Name:'PpEffectRandom'; value:$00000201),
    (Name:'PpEffectBlindsHorizontal'; value:$00000301),
    (Name:'PpEffectBlindsVertical'; value:$00000302),
    (Name:'PpEffectCheckerboardAcross'; value:$00000401),
    (Name:'PpEffectCheckerboardDown'; value:$00000402),
    (Name:'PpEffectCoverLeft'; value:$00000501),
    (Name:'PpEffectCoverUp'; value:$00000502),
    (Name:'PpEffectCoverRight'; value:$00000503),
    (Name:'PpEffectCoverDown'; value:$00000504),
    (Name:'PpEffectCoverLeftUp'; value:$00000505),
    (Name:'PpEffectCoverRightUp'; value:$00000506),
    (Name:'PpEffectCoverLeftDown'; value:$00000507),
    (Name:'PpEffectCoverRightDown'; value:$00000508),
    (Name:'PpEffectDissolve'; value:$00000601),
    (Name:'PpEffectFade'; value:$00000701),
    (Name:'PpEffectUncoverLeft'; value:$00000801),
    (Name:'PpEffectUncoverUp'; value:$00000802),
    (Name:'PpEffectUncoverRight'; value:$00000803),
    (Name:'PpEffectUncoverDown'; value:$00000804),
    (Name:'PpEffectUncoverLeftUp'; value:$00000805),
    (Name:'PpEffectUncoverRightUp'; value:$00000806),
    (Name:'PpEffectUncoverLeftDown'; value:$00000807),
    (Name:'PpEffectUncoverRightDown'; value:$00000808),
    (Name:'PpEffectRandomBarsHorizontal'; value:$00000901),
    (Name:'PpEffectRandomBarsVertical'; value:$00000902),
    (Name:'PpEffectStripsUpLeft'; value:$00000A01),
    (Name:'PpEffectStripsUpRight'; value:$00000A02),
    (Name:'PpEffectStripsDownLeft'; value:$00000A03),
    (Name:'PpEffectStripsDownRight'; value:$00000A04),
    (Name:'PpEffectStripsLeftUp'; value:$00000A05),
    (Name:'PpEffectStripsRightUp'; value:$00000A06),
    (Name:'PpEffectStripsLeftDown'; value:$00000A07),
    (Name:'PpEffectStripsRightDown'; value:$00000A08),
    (Name:'PpEffectWipeLeft'; value:$00000B01),
    (Name:'PpEffectWipeUp'; value:$00000B02),
    (Name:'PpEffectWipeRight'; value:$00000B03),
    (Name:'PpEffectWipeDown'; value:$00000B04),
    (Name:'PpEffectBoxOut'; value:$00000C01),
    (Name:'PpEffectBoxIn'; value:$00000C02),
    (Name:'PpEffectFlyFromLeft'; value:$00000D01),
    (Name:'PpEffectFlyFromTop'; value:$00000D02),
    (Name:'PpEffectFlyFromRight'; value:$00000D03),
    (Name:'PpEffectFlyFromBottom'; value:$00000D04),
    (Name:'PpEffectFlyFromTopLeft'; value:$00000D05),
    (Name:'PpEffectFlyFromTopRight'; value:$00000D06),
    (Name:'PpEffectFlyFromBottomLeft'; value:$00000D07),
    (Name:'PpEffectFlyFromBottomRight'; value:$00000D08),
    (Name:'PpEffectPeekFromLeft'; value:$00000D09),
    (Name:'PpEffectPeekFromDown'; value:$00000D0A),
    (Name:'PpEffectPeekFromRight'; value:$00000D0B),
    (Name:'PpEffectPeekFromUp'; value:$00000D0C),
    (Name:'PpEffectCrawlFromLeft'; value:$00000D0D),
    (Name:'PpEffectCrawlFromUp'; value:$00000D0E),
    (Name:'PpEffectCrawlFromRight'; value:$00000D0F),
    (Name:'PpEffectCrawlFromDown'; value:$00000D10),
    (Name:'PpEffectZoomIn'; value:$00000D11),
    (Name:'PpEffectZoomInSlightly'; value:$00000D12),
    (Name:'PpEffectZoomOut'; value:$00000D13),
    (Name:'PpEffectZoomOutSlightly'; value:$00000D14),
    (Name:'PpEffectZoomCenter'; value:$00000D15),
    (Name:'PpEffectZoomBottom'; value:$00000D16),
    (Name:'PpEffectStretchAcross'; value:$00000D17),
    (Name:'PpEffectStretchLeft'; value:$00000D18),
    (Name:'PpEffectStretchUp'; value:$00000D19),
    (Name:'PpEffectStretchRight'; value:$00000D1A),
    (Name:'PpEffectStretchDown'; value:$00000D1B),
    (Name:'PpEffectSwivel'; value:$00000D1C),
    (Name:'PpEffectSpiral'; value:$00000D1D),
    (Name:'PpEffectSplitHorizontalOut'; value:$00000E01),
    (Name:'PpEffectSplitHorizontalIn'; value:$00000E02),
    (Name:'PpEffectSplitVerticalOut'; value:$00000E03),
    (Name:'PpEffectSplitVerticalIn'; value:$00000E04),
    (Name:'PpEffectFlashOnceFast'; value:$00000F01),
    (Name:'PpEffectFlashOnceMedium'; value:$00000F02),
    (Name:'PpEffectFlashOnceSlow'; value:$00000F03),
    (Name:'PpEffectAppear'; value:$00000F04)
  );
begin
  Result := GetOleEnumText(Data, 80, AValue);
end;

function PpFollowColorsStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'PpFollowColorsMixed'; Value:$FFFFFFFE),
    (Name:'PpFollowColorsNone'; Value:$00000000),
    (Name:'PpFollowColorsScheme'; Value:$00000001),
    (Name:'PpFollowColorsTextAndBackground'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function PpMediaTypeStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'PpMediaTypeMixed'; Value:$FFFFFFFE),
    (Name:'PpMediaTypeOther'; Value:$00000001),
    (Name:'PpMediaTypeSound'; Value:$00000002),
    (Name:'PpMediaTypeMovie'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function PpPlaceholderTypeStr(AValue: LongWord): string;
const
  Data : array[0..$10] of TOleEnumItem = (
    (Name:'PpPlaceholderMixed'; Value:$FFFFFFFE),
    (Name:'PpPlaceholderTitle'; Value:$00000001),
    (Name:'PpPlaceholderBody'; Value:$00000002),
    (Name:'PpPlaceholderCenterTitle'; Value:$00000003),
    (Name:'PpPlaceholderSubtitle'; Value:$00000004),
    (Name:'PpPlaceholderVerticalTitle'; Value:$00000005),
    (Name:'PpPlaceholderVerticalBody'; Value:$00000006),
    (Name:'PpPlaceholderObject'; Value:$00000007),
    (Name:'PpPlaceholderChart'; Value:$00000008),
    (Name:'PpPlaceholderBitmap'; Value:$00000009),
    (Name:'PpPlaceholderMediaClip'; Value:$0000000A),
    (Name:'PpPlaceholderOrgChart'; Value:$0000000B),
    (Name:'PpPlaceholderTable'; Value:$0000000C),
    (Name:'PpPlaceholderSlideNumber'; Value:$0000000D),
    (Name:'PpPlaceholderHeader'; Value:$0000000E),
    (Name:'PpPlaceholderFooter'; Value:$0000000F),
    (Name:'PpPlaceholderDate'; Value:$00000010)
  );
begin
  Result := GetOleEnumText(Data, $10+1, AValue);
end;

function PpSlideLayoutStr(AValue: LongWord): string;
const
  Data : array[0..$1C] of TOleEnumItem = (
    (Name:'PpLayoutMixed'; Value:$FFFFFFFE),
    (Name:'PpLayoutTitle'; Value:$00000001),
    (Name:'PpLayoutText'; Value:$00000002),
    (Name:'PpLayoutTwoColumnText'; Value:$00000003),
    (Name:'PpLayoutTable'; Value:$00000004),
    (Name:'PpLayoutTextAndChart'; Value:$00000005),
    (Name:'PpLayoutChartAndText'; Value:$00000006),
    (Name:'PpLayoutOrgchart'; Value:$00000007),
    (Name:'PpLayoutChart'; Value:$00000008),
    (Name:'PpLayoutTextAndClipart'; Value:$00000009),
    (Name:'PpLayoutClipartAndText'; Value:$0000000A),
    (Name:'PpLayoutTitleOnly'; Value:$0000000B),
    (Name:'PpLayoutBlank'; Value:$0000000C),
    (Name:'PpLayoutTextAndObject'; Value:$0000000D),
    (Name:'PpLayoutObjectAndText'; Value:$0000000E),
    (Name:'PpLayoutLargeObject'; Value:$0000000F),
    (Name:'PpLayoutObject'; Value:$00000010),
    (Name:'PpLayoutTextAndMediaClip'; Value:$00000011),
    (Name:'PpLayoutMediaClipAndText'; Value:$00000012),
    (Name:'PpLayoutObjectOverText'; Value:$00000013),
    (Name:'PpLayoutTextOverObject'; Value:$00000014),
    (Name:'PpLayoutTextAndTwoObjects'; Value:$00000015),
    (Name:'PpLayoutTwoObjectsAndText'; Value:$00000016),
    (Name:'PpLayoutTwoObjectsOverText'; Value:$00000017),
    (Name:'PpLayoutFourObjects'; Value:$00000018),
    (Name:'PpLayoutVerticalText'; Value:$00000019),
    (Name:'PpLayoutClipArtAndVerticalText'; Value:$0000001A),
    (Name:'PpLayoutVerticalTitleAndText'; Value:$0000001B),
    (Name:'PpLayoutVerticalTitleAndTextOverChart'; Value:$0000001C)
  );
begin
  Result := GetOleEnumText(Data, $1C+1, AValue);
end;

function PpSlideShowRangeTypeStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'PpShowAll'; Value:$00000001),
    (Name:'PpShowSlideRange'; Value:$00000002),
    (Name:'PpShowNamedSlideShow'; Value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;

function PpSlideShowTypeStr(AValue: LongWord): string;
const
  Data : array[1..3] of TOleEnumItem = (
    (Name:'PpShowTypeSpeaker'; Value:$00000001),
    (Name:'PpShowTypeWindow'; Value:$00000002),
    (Name:'PpShowTypeKiosk'; Value:$00000003)
  );
begin
  result := GetOleEnumText(Data, 3, AValue);
end;

function PpSlideSizeTypeStr(AValue: LongWord): string;
const
  Data : array[1..7] of TOleEnumItem = (
    (Name:'PpSlideSizeOnScreen'; Value:$00000001),
    (Name:'PpSlideSizeLetterPaper'; Value:$00000002),
    (Name:'PpSlideSizeA4Paper'; Value:$00000003),
    (Name:'PpSlideSize35MM'; Value:$00000004),
    (Name:'PpSlideSizeOverhead'; Value:$00000005),
    (Name:'PpSlideSizeBanner'; Value:$00000006),
    (Name:'PpSlideSizeCustom'; Value:$00000007)
  );
begin
  Result := GetOleEnumText(Data, 7, AValue);
end;

function PpSoundEffectTypeStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'PpSoundEffectsMixed'; Value:$FFFFFFFE),
    (Name:'PpSoundNone'; Value:$00000000),
    (Name:'PpSoundStopPrevious'; Value:$00000001),
    (Name:'PpSoundFile'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function PpSoundFormatTypeStr(AValue: LongWord): string;
const
  Data : array[0..4] of TOleEnumItem = (
    (Name:'PpSoundFormatMixed'; value:$FFFFFFFE),
    (Name:'PpSoundFormatNone'; value:$00000000),
    (Name:'PpSoundFormatWAV'; value:$00000001),
    (Name:'PpSoundFormatMIDI'; value:$00000002),
    (Name:'PpSoundFormatCDAudio'; value:$00000003)
  );
begin
  Result := GetOleEnumText(Data, 5, AValue);
end;

function PpTextLevelEffectStr(AValue: LongWord): string;
const
  Data : array[0..7] of TOleEnumItem = (
    (Name:'PpAnimateLevelMixed'; Value:$FFFFFFFE),
    (Name:'PpAnimateLevelNone'; Value:$00000000),
    (Name:'PpAnimateByFirstLevel'; Value:$00000001),
    (Name:'PpAnimateBySecondLevel'; Value:$00000002),
    (Name:'PpAnimateByThirdLevel'; Value:$00000003),
    (Name:'PpAnimateByFourthLevel'; Value:$00000004),
    (Name:'PpAnimateByFifthLevel'; Value:$00000005),
    (Name:'PpAnimateByAllLevels'; Value:$00000010)
  );
begin
  Result := GetOleEnumText(Data, 8, AValue);
end;

function PpTextUnitEffectStr(AValue: LongWord): string;
const
  Data : array[0..3] of TOleEnumItem = (
    (Name:'PpAnimateUnitMixed'; Value:$FFFFFFFE),
    (Name:'PpAnimateByParagraph'; Value:$00000000),
    (Name:'PpAnimateByWord'; Value:$00000001),
    (Name:'PpAnimateByCharacter'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 4, AValue);
end;

function PpUpdateOptionStr(AValue: LongWord): string;
const
  Data : array[0..2] of TOleEnumItem = (
    (Name:'PpUpdateOptionMixed'; Value:$FFFFFFFE),
    (Name:'PpUpdateOptionManual'; Value:$00000001),
    (Name:'PpUpdateOptionAutomatic'; Value:$00000002)
  );
begin
  Result := GetOleEnumText(Data, 3, AValue);
end;


end.

