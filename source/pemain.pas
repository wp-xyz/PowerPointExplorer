unit peMain;

{$mode Delphi}

interface

uses
  Windows, Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ActnList, Menus,
  ComCtrls, ExtCtrls, StdCtrls, mrumanager;

type

  { TMainForm }

  TMainForm = class(TForm)
    acFileOpen: TAction;
    acFileExit: TAction;
    acFileSaveAs: TAction;
    acEditCopy: TAction;
    acEditLink: TAction;
    acViewCollapseAll: TAction;
    acViewExpandAll: TAction;
    ActionList: TActionList;
    Bevel1: TBevel;
    CbShowErrorItems: TCheckBox;
    ImageList: TImageList;
    ImgReadOnly: TImage;
    ImgReadWrite: TImage;
    Label1: TLabel;
    LblReadOnly: TLabel;
    LblReadWrite: TLabel;
    MenuItem2: TMenuItem;
    mnuEditLink: TMenuItem;
    mnuEditCopy: TMenuItem;
    mnuEdit: TMenuItem;
    mnuViewCollapseAll: TMenuItem;
    mnuViewExpandAll: TMenuItem;
    mnuView: TMenuItem;
    mnuFileSaveAs: TMenuItem;
    OpenDialog: TOpenDialog;
    Panel4: TPanel;
    SaveDialog: TSaveDialog;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    TreeViewImages: TImageList;
    ValueList: TListView;
    MainMenu: TMainMenu;
    MenuItem1: TMenuItem;
    mnuOpenRecent: TMenuItem;
    mnuFile: TMenuItem;
    mnuFileOpen: TMenuItem;
    mnuFileExit: TMenuItem;
    MRUMenuManager: TMRUMenuManager;
    Panel1: TPanel;
    Panel2: TPanel;
    RecentFilesPopup: TPopupMenu;
    Splitter1: TSplitter;
    ToolBar1: TToolBar;
    tbFileOpen: TToolButton;
    tbFileExit: TToolButton;
    TreeView: TTreeView;
    procedure acEditCopyExecute(Sender: TObject);
    procedure acEditLinkExecute(Sender: TObject);
    procedure acFileExitExecute(Sender: TObject);
    procedure acFileOpenExecute(Sender: TObject);
    procedure acFileSaveAsExecute(Sender: TObject);
    procedure ActionListUpdate(AAction: TBasicAction; var Handled: Boolean);
    procedure acViewCollapseAllExecute(Sender: TObject);
    procedure acViewExpandAllExecute(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MRUMenuManagerRecentFile(Sender: TObject; const AFileName: String);
    procedure TreeViewChange(Sender: TObject; Node: TTreeNode);
    procedure TreeViewExpanding(Sender: TObject; ANode: TTreeNode;
      var AllowExpansion: Boolean);
    procedure ValueListDblClick(Sender: TObject);
  private
    FFileName: String;
    FActSlideID: integer;
    FActPresentation: OLEVariant;
    FPowerpoint: OLEVariant;
    FPowerpointRunning: boolean;
    FPowerpointLeft, FPowerpointTop, FPowerPointHeight, FPowerPointWidth : integer;
    FShowErrorItems: Boolean;
    FTestPresentation : OLEVariant;

    procedure AddPresentation(const AFileName: String);

    procedure DisplayActionSettings(ASettings:OLEVariant; AIndex:integer);
    procedure DisplayAnimationBehavior(AValue: OLEVariant);
    procedure DisplayAnimationSettings(ASettings: OLEVariant);
    procedure DisplayCallout(ACallout: OLEVariant);
    procedure DisplayColor(const Prefix: string; AColor: OLEVariant; AReadOnly: integer);
    procedure DisplayColorEffect(const Prefix: string; AEffect: OLEVariant);
    procedure DisplayCreator(ACreator: OLEVariant);
    procedure DisplayDiagram(ADiagram: OLEVariant);
    procedure DisplayDocProperties(AProps: OLEVariant);
    procedure DisplayEffect(AEffect: OLEVariant);
    procedure DisplayEffectInformation(AInfo: OLEVariant);
    procedure DisplayEffectParameters(const Prefix: string; AParam: OLEVariant);
    procedure DisplayError(const ACaption, AMsg: string);
    procedure DisplayFillFormat(AFillFormat :OLEVariant);
    procedure DisplayGroupItems(AGroupItems: OLEVariant);
    procedure DisplayHeaderFooter(AHeaderFooter: OLEVariant);
    procedure DisplayHeadersFooters(AHeadersFooters: OLEVariant);
    procedure DisplayHyperlinks(AHyperlinks: OLEVariant);
    procedure DisplayLineFormat(AFormat: OLEVariant);
    procedure DisplayLinkFormat(AFormat: OLEVariant);
    procedure DisplayMaster(AMaster: OLEVariant);
    procedure DisplayMotionEffect(const Prefix: string; AEffect: OLEVariant);
    procedure DisplayOLEFormat(AOLEFormat: OLEVariant);
    procedure DisplayPageSetup(APageSetup: OLEVariant);
    procedure DisplayPictureFormat(APicFormat: OLEVariant);
    procedure DisplayPlaySettings(const Prefix: string; ASettings: OLEVariant);
    procedure DisplayPoints(const ACaption: string; APoints: OLEVariant; AReadOnly: integer);
    procedure DisplayPresentation(APresentation: OLEVariant);
    procedure DisplayPropertyEffect(const Prefix: string; AEffect: OLEVariant);
    procedure DisplayRGB(const Prefix: string; RGB: OLEVariant; AReadOnly: integer);
    procedure DisplayRotationEffect(const Prefix: string; AEffect: OLEVariant);
    procedure DisplayScaleEffect(const Prefix: string; AEffect: OLEVariant);
    procedure DisplaySequence(ASequence: OLEVariant);
    procedure DisplayShadow(AShadow: OLEVariant);
    procedure DisplayShape(AShape: OLEVariant);
    procedure DisplayShapes(AShapes: OLEVariant);
    procedure DisplaySlide(ASlide: OLEVariant);
    procedure DisplaySlides(ASlides: OLEVariant);
    procedure DisplaySlideshowSettings(ASettings: OLEVariant);
    procedure DisplaySlideshowTransition(ASettings: OLEVariant);
    procedure DisplaySoundEffect(AEffect: OLEVariant);
    procedure DisplayTextEffect(ATextEffect: OLEVariant);
    procedure DisplayTextFrame(ATextFrame: OLEVariant);
    procedure DisplayTextRange(const PreFix: string; ATextRange: OLEVariant);
    procedure DisplayTiming(const Prefix: string; AValue: OLEVariant);
    procedure DisplayValue(const AName: string; AValue: variant;
      AReadOnly: integer; ADisplayOption: char = #0; AEditOption: char = #0);

    function FindPresentation(ANode: TTreeNode): OLEVariant;
    function FindPresentationNode(ANode: TTreeNode): TTreeNode;
    function FindShape(ANode:TTreeNode): OLEVariant; overload;
    function FindShape(AShapeCollection: OLEVariant; const AShapeName: string): OLEVariant; overload;
    function FindSlide(ANode: TTreeNode): OLEVariant; overload;
    function FindSlide(ASlideName: string): OLEVariant; overload;

    function IsDirty(ANode: TTreeNode): Boolean;
    function NodeIsShape(ANode: TTreeNode): Boolean;
    function NodeIsSlide(ANode: TTreeNode): Boolean;
    procedure PopulateTreeView(APresentationIndex: Integer);
    procedure SetDirty(ANode: TTreeNode; AValue: boolean);

    function SelectEffect(ANode: TTreeNode): OLEVariant;
    function SelectHeaderFooter(ANode: TTreeNode): OLEVariant;
    function SelectMaster(ANode: TTreeNode): OLEVariant;
    function SelectPresentation(ANode: TTreeNode): OLEVariant;
    function SelectSequence(ANode: TTreeNode): OLEVariant;
    function SelectShape(ANode: TTreeNode): OLEVariant;
    function SelectSlide(ANode: TTreeNode): OLEVariant;

    procedure ReadIni;
    procedure WriteIni;

  public
    property PowerPoint: OLEVariant read FPowerPoint;

  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

uses
  FileUtil, IniFiles, Variants, ComObj,
  peOffice, peUtils, pePropertyViewer;

function GetRGBName(AColor: Integer): string;
type
  TCol = packed record R,G,B,X: byte end;
begin
  Result := Format('R=%d G=%d B=%d', [TCol(AColor).R, TCol(AColor).G, TCol(AColor).B]);
end;


{ TMainForm }

procedure TMainForm.acFileExitExecute(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.acEditCopyExecute(Sender: TObject);
var
  node: TTreeNode;
  slide: OLEVariant;
  shape: OLEVariant;
begin
  node := TreeView.Selected;
  if NodeIsSlide(Node) then
  begin
    slide := FindSlide(Node);
    slide.Copy;
  end else
  if NodeIsShape(Node) then
  begin
    shape := FindShape(Node);
    shape.Copy;
  end;
end;

procedure TMainForm.acEditLinkExecute(Sender: TObject);
// Edit links of some types
// Little error checking. It works or not, no harm if not...
// Ref.: http://pptfaq.com/FAQ00433.htm
var
  oldLinkSource: string;
  newLinkSource: string;
  newLinkPath: string;
  shape: OLEVariant;
begin
  shape := FindShape(Treeview.Selected);
  try
    oldLinkSource := shape.Linkformat.SourceFullName;
    newLinkSource := InputBox('Link editor', 'Edit link', oldLinkSource);
    if newLinkSource = oldLinkSource then // nothing changed --> our work is done.
      exit;

    if NewLinkSource = '' then // the user cancelled --> quit
      exit;

    // Is it a valid filename? Is the file where it belongs?
    // Test against the filename portion of the link in case the link includes
    // range information
    if IsRelativePath(newLinkSource) then
      newLinkPath := FActPresentation.Path + '\' + newLinkSource
    else
      newLinkPath := newLinkSource;
    if FileExists(newLinkPath) then
    begin
      shape.LinkFormat.SourceFullName := newLinkSource;
      try
        shape.LinkFormat.Update;
      except
      end;
      SetDirty(Treeview.Selected, true);
    end else begin
      MessageDlg(Format(
        'File "%s" not found. Please copy the previousely linked '+
        'file to its new location before attempting to alter the link, '+
        'then try again.',
        [newLinkSource]),
        mtInformation, [mbOK], 0
      );
    end;
  except
    MessageDlg('Please select a PowerPoint shape object with an editable link, then try again.',
      mtInformation, [mbOK], 0);
  end;
end;

procedure TMainForm.acFileOpenExecute(Sender: TObject);
begin
  with OpenDialog do
  begin
    if FileName <> '' then
      InitialDir := ExtractFileDir(FileName);
    FileName := '';
    if Execute then begin
      AddPresentation(FileName);
      MRUMenuManager.AddToRecent(FileName);
    end;
  end;
end;

procedure TMainForm.acFileSaveAsExecute(Sender: TObject);
begin
  if TreeView.Selected = nil then
    exit;

  with SaveDialog do
  begin
    FileName := FActPresentation.Filename;
    InitialDir := ExtractFileDir(Filename);
    if Execute then
    begin
      FActPresentation.SaveCopyAs(FileName);
      SetDirty(Treeview.Selected, false);
    end;
  end;
end;

procedure TMainForm.ActionListUpdate(AAction: TBasicAction;
  var Handled: Boolean);
var
  node: TTreeNode;
begin
  node := TreeView.Selected;
  AcViewExpandAll.Enabled := TreeView.Items.Count > 0;
  AcViewCollapseAll.Enabled := AcViewExpandAll.Enabled;
  AcEditCopy.Enabled := NodeIsShape(Node) or NodeIsSlide(Node);
//  AcClearRecentFiles.Enabled := RecentFiles.RecentCount > 0;
  AcFileSaveAs.Enabled := node <> nil;
  Handled := true;
end;

procedure TMainForm.acViewCollapseAllExecute(Sender: TObject);
begin
  TreeView.FullCollapse;
end;

procedure TMainForm.acViewExpandAllExecute(Sender: TObject);
begin
  TreeView.BeginUpdate;
  try
    TreeView.FullExpand;
  finally
    TreeView.EndUpdate;
  end;
end;

procedure TMainForm.AddPresentation(const AFileName: string);
var
  crs: TCursor;
  i: integer;
  N: TTreeNode;
begin
  crs := Screen.Cursor;
  Screen.Cursor := crHourglass;
  try
    for i:=1 to FPowerpoint.Presentations.Count do
    begin
      if CompareText(FPowerpoint.Presentations.Item(i).FullName, AFileName) = 0 then
      begin
        FFileName := AFileName;
        FActPresentation := FPowerPoint.Presentations.Item(i);
        if TreeView.Items.Count > 0 then
        begin
          N := TreeView.Items[0];
          while Assigned(N) do begin
            if CompareText(ExtractFileName(AFileName), N.Text) = 0 then
            begin
              TreeView.Selected := N;
              exit;
            end;
            N := N.GetNextSibling;
          end;
        end;
      end;
    end;

    FFileName := AFilename;
    try
      FActPresentation := FPowerpoint.Presentations.Open(WideString(FFileName), msoTrue, msoFalse, msoTrue);
    except
      FActPresentation := FPowerpoint.Presentations.Open(WideString(FFileName), msoTrue, msoFalse, msoFalse);
      // ppt could not be opened without error. Try again without window (last parameter).
    end;

    PopulateTreeView(FPowerpoint.Presentations.Count);
    SetDirty(Treeview.Selected, false);
  finally
    Screen.Cursor := crs;
  end;
end;

procedure TMainForm.DisplayActionSettings(ASettings: OLEVariant; AIndex: Integer);
var
  settings: OLEVariant;
begin
  if VarIsEmpty(ASettings) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      settings := ASettings.Item(AIndex);

      try
        DisplayValue('Action', PpActionTypeStr(settings.Action), RW);
      except
        on E:Exception do DisplayError('Action', E.Message);
      end;

      try
        DisplayValue('ActionVerb', settings.ActionVerb, RW);
      except
        on E:Exception do DisplayError('ActionVerb', E.Message);
      end;

      try
        DisplayValue('AnimateAction', MsoTriStateStr(settings.AnimateAction), RW);
      except
        on E:Exception do DisplayError('AnimateAction', E.Message);
      end;

      try
        DisplayValue('Hyperlink', settings.Hyperlink.Address, RO);
      except
        on E:Exception do DisplayError('Hyperlink', E.Message);
      end;

      try
        DisplaySoundEffect(settings.SoundEffect);
      except
        on E:Exception do DisplayError('SoundEffect', E.Message);
      end;

    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayAnimationBehavior(AValue: OLEVariant);
begin
  if VarIsEmpty(AValue) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      DisplayValue('Accumulate', MsoAnimAccumulateStr(AValue.Accumulate), RW);
      DisplayValue('Additive', MsoAnimAdditiveStr(AValue.Additive), RW);

      try
        DisplayColorEffect('ColorEffect', AValue.ColorEffect);
      except
        on E:Exception do
          DisplayError('ColorEffect', E.Message);
      end;

      try
        DisplayMotionEffect('MotionEffect', AValue.MotionEffect);
      except
        on E:Exception do
          DisplayError('MotionEffect', E.Message);
      end;

      try
        DisplayPropertyEffect('PropertyEffect', AValue.PropertyEffect);
      except
        on E:Exception do
          DisplayError('PropertyEffect', E.Message);
      end;

      try
        DisplayRotationEffect('RotationEffect', AValue.RotationEffect);
      except
        on E:Exception do
          DisplayError('RotationEffect', E.Message);
      end;

      try
        DisplayScaleEffect('ScaleEffect', AValue.ScaleEffect);
      except
        on E:Exception do
          DisplayError('ScaleEffect', E.Message);
      end;

      try
        DisplayTiming('Timing', AValue.Timing);
      except
        on E:Exception do DisplayError('Timing', E.Message);
      end;

      DisplayValue('Type', MsoAnimTypeStr(AValue.&Type), RW);
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayAnimationSettings(ASettings:OLEVariant);
var
  w : longword;
begin
  if VarIsEmpty(ASettings) then exit;
  with ValueList do begin
    Items.BeginUpdate;
    try
      Items.Clear;
      w := ASettings.Animate;
      Displayvalue('AdvanceMode', PpAdvanceModeStr(ASettings.AdvanceMode), RW);
      DisplayValue('AdvanceTime', ASettings.AdvanceTime, RW);
      DisplayValue('AfterEffect', PpAfterEffectStr(ASettings.AfterEffect), RW);
      DisplayValue('Animate', MsoTriStateStr(ASettings.Animate), RW);
      DisplayValue('AnimateBackground', MsoTriStateStr(ASettings.AnimateBackground), RW);
      DisplayValue('AnimateTextInReverse', MsoTriStateStr(ASettings.AnimateTextInReverse), RW);
      if w=msoTrue
        then DisplayValue('AnimationOrder', ASettings.AnimationOrder, RW)
        else DisplayValue('AnimationOrder', '- Wert nicht bestimmbar -', RW);
      DisplayValue('ChartUnitEffect', PpChartUnitEffectStr(ASettings.ChartUnitEffect), RW);
      DisplayColor('DimColor', ASettings.DimColor, RW);
      DisplayValue('EntryEffect', PpEntryEffectStr(ASettings.EntryEffect), RW);
      try
        DisplayPlaySettings('PlaySettings', ASettings.PlaySettings);
      except
        on E:Exception do DisplayError('PlaySettings', E.Message);
      end;
      try
        DisplaySoundEffect(ASettings.SoundEffect);
      except
        on E:Exception do DisplayError('SoundEffect', E.Message);
      end;
      Displayvalue('TextLevelEffect', PpTextLevelEffectStr(ASettings.TextLevelEffect), RW);
      DisplayValue('TextUnitEffect', PpTextUnitEffectStr(ASettings.TextUnitEffect), RW);
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayCallout(ACallout: OLEVariant);
begin
  if VarIsEmpty(ACallout) then
    exit;
  with ValueList do begin
    Items.BeginUpdate;
    try
      Items.Clear;
      try
        DisplayValue('Accent', MsoTriStateStr(ACallout.Accent), RW);
      except
        on E:Exception do
          DisplayError('Accent', E.Message);
      end;

      try
        DisplayValue('Angle', MsoCalloutAngleTypeStr(ACallout.Angle), RW);
      except
        on E:Exception do
          DisplayError('Angle', E.Message);
      end;

      try
        DisplayValue('AutoAttach', MsoTriStateStr(ACallout.AutoAttach), RW);
      except
        on E:Exception do
          DisplayError('AutoAttach', E.Message);
      end;

      try
        DisplayValue('AutoLength', MsoTriStateStr(ACallout.AutoLength), RW);
      except
        on E:Exception do
          DisplayError('AutoLength', E.Message);
      end;

      try
        DisplayValue('Border', MsoTriStateStr(ACallout.Border), RW);
      except
        on E:Exception do
          DisplayError('Border', E.Message);
      end;

      DisplayCreator(ACallout.Creator);
      try
        DisplayValue('Drop', ACallout.Drop, RO);
      except
        on E:Exception do
          DisplayError('Drop', E.Message);
      end;

      try
        DisplayValue('DropType',
          MsoCalloutDropTypeStr(ACallout.DropType), RO);
      except
        on E:Exception do
          DisplayError('DropType', E.Message);
      end;

      try
        DisplayValue('Gap', ACallout.Gap, RW);
      except
        on E:Exception do
          DisplayError('Gap', E.Message);
      end;

      try
        DisplayValue('Length', ACallout.Length, RO);
      except
        on E:Exception do
          DisplayError('Length', E.Message);
      end;

      try
        DisplayValue('Type', MsoCalloutTypeStr(ACallout.&Type), RW);
      except
        on E:Exception do
          DisplayError('Type', E.Message);
      end;

    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayColor(const Prefix: string; AColor: OLEVariant;
  AReadOnly: Integer);
var
  z: double;
begin
  DisplayRGB(Prefix, AColor.RGB, AReadOnly);

  try
    DisplayValue(Prefix+'.SchemeColor', PpColorSchemeIndexStr(AColor.SchemeColor),
      AReadOnly);
  except
    on E:Exception do
      DisplayError(Prefix+'.SchemeColor', E.Message);
  end;

  try
    DisplayValue(Prefix+'.TintAndShade', AColor.TintAndShade, AReadOnly);
  except
    on E:Exception do
      DisplayError(Prefix+'.TintAndShade', E.Message);
  end;

{  DisplayValue(Prefix+'y', MsoColorTypeStr(AColor.Type), RO);}
end;

procedure TMainForm.DisplayColorEffect(const Prefix: string; AEffect: OLEVariant);
begin
  try
    DisplayRGB(Prefix+'.By', AEffect.By, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.By', E.Message);
  end;
  try
    DisplayRGB(Prefix+'.From', AEffect.From, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.From', E.Message);
  end;
  try
    DisplayRGB(Prefix+'.To', AEffect.&To, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.To', E.Message);
  end;
end;

procedure TMainForm.DisplayCreator(ACreator: OLEVariant);
var
  n: integer;
  s: string;
begin
  n := ACreator;
  s := Format('$%x', [n]);;
  if n = $50575054 then s := s + ' (Powerpoint)';
  DisplayValue('Creator', s, RO);
end;

procedure TMainForm.DisplayDiagram(ADiagram: OLEVariant);
begin
  if VarIsEmpty(ADiagram) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      DisplayCreator(ADiagram.Creator);
      DisplayValue('AutoFormat', MsoTriStateStr(ADiagram.AutoFormat), RW);
      DisplayValue('AutoLayout', MsoTriStateStr(ADiagram.AutoLayout), RW);
      DisplayValue('Reverse', MsoTriSTateStr(ADiagram.Reverse), RW);
      {DisplayValue('Type', MsoDiagramTypeStr(ADiagram.Type), RO);}
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayDocProperties(AProps: OLEVariant);
var
  i: integer;
begin
  if VarIsEmpty(AProps) then
    exit;

  with ValueList do
  begin
    Items.Clear;
    for i:=1 to AProps.Count do
    begin
      try
        DisplayValue(AProps.Item[i].Name, AProps.Item[i].Value, RO);
      except
        on E:Exception do
          DisplayError(AProps.Item[i].Name, E.Message);
      end;
    end;
  end;
end;

procedure TMainForm.DisplayEffect(AEffect: OLEVariant);
begin
  if VarIsEmpty(AEffect) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      DisplayValue('Behaviors.Count', AEffect.Behaviors.Count, RO);
      DisplayValue('DisplayName', AEffect.DisplayName, RO);
      DisplayEffectParameters('EffectParameters', AEffect.EffectParameters);
      DisplayValue('EffectType', MsoAnimEffectStr(AEffect.EffectType), RW);
      DisplayValue('Exit', MsoTriStateStr(AEffect.Exit), RW);
      DisplayValue('Index', AEffect.Index, RO);

      try
        DisplayValue('Paragraph', AEffect.Paragraph, RW);
      except
        on E:Exception do DisplayError('Paragraph', E.Message);
      end;

      DisplayValue('Shape.Name', AEffect.Shape.Name, RO);

      try
        DisplayValue('TextRangeLength', AEffect.TextRangeLength, RO);
      except
        on E:Exception do DisplayError('TextRangeLength', E.Message);
      end;

      try
        DisplayValue('TextRangeStart', AEffect.TextRangeStart, RO);
      except
        on E:Exception do DisplayError('TextRangeStart', E.Message);
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayEffectInformation(AInfo: OLEVariant);
begin
  if VarIsEmpty(AInfo) then
    exit;
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      DisplayValue('AfterEffect', MsoAnimAfterEffectStr(AInfo.AfterEffect), RO);
      DisplayValue('AnimateBackground', MsoTriStateStr(AInfo.AnimateBackground), RW);
      DisplayValue('AnimateTextInReverse', MsoTriStateStr(AInfo.AnimateTextInReverse), RW);

      try
        DisplayValue('BuldByLevelEffect', MsoAnimateByLevelStr(AInfo.BuildByLevelEffect), RO);
      except
        on E:Exception do DisplayError('BuildByLevelEffect', E.Message);
      end;

      DisplayColor('Dim', AInfo.Dim, RO);
      DisplayPlaySettings('PlaySettings', AInfo.PlaySettings);
      DisplaySoundEffect(AInfo.SoundEffect);
      DisplayValue('TextUnitEffect', MsoAnimTextUnitEffectStr(AInfo.TextUnitEffect), RO);
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayEffectParameters(const Prefix: string;
  AParam: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.Amount', AParam.Amount, RW);
  except
    on E:Exception do DisplayError(Prefix+'.Amount', E.Message);
  end;

  try
    DisplayColor(Prefix+'.Color1', AParam.Color1, RW);
  except
    on E:Exception do DisplayError(Prefix+'.Color1', E.Message);
  end;

  try
    DisplayColor(Prefix+'.Color2', AParam.Color2, RW);
  except
    on E:Exception do DisplayError(Prefix+'.Color2', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Direction', MsoAnimDirectionStr(AParam.Direction), RW);
  except
    on E:Exception do DisplayError(Prefix+'.Direction', E.Message);
  end;

  try
    DisplayValue(Prefix+'.FontName', AParam.FontName, RW);
  except
    on E:Exception do DisplayError(Prefix+'.FontName', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Relative', MsoTriStateStr(AParam.Relative), RW);
  except
    on E:Exception do DisplayError(Prefix+'.Relative', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Size', AParam.Size, RW);
  except
    on E:Exception do DisplayError(Prefix+'.Size', E.Message);
  end;
end;

procedure TMainForm.DisplayError(const ACaption, AMsg:string);
begin
  if FShowErrorItems then
    DisplayValue(ACaption, Format('--- %s ---', [AMsg]), RO);
end;

procedure TMainForm.DisplayFillFormat(AFillFormat: OLEVariant);
begin
  if VarIsEmpty(AFillFormat) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      DisplayCreator(AFillFormat.Creator);
      DisplayColor('BackColor', AFillFormat.BackColor, RW);
      DisplayColor('ForeColor', AFillFormat.ForeColor, RW);
      DisplayValue('GradientColorType', MsoGradientColorTypeStr(AFillFormat.GradientColorType), RO);

      try
        DisplayValue('GradientDegree', AFillFormat.GradientDegree, RO);
      except
        on E:Exception do
          DisplayError('GradientDegree', E.Message);
      end;

      DisplayValue('GradientStyle', MsoGradientStyleStr(AFillFormat.GradientStyle), RO);
      DisplayValue('GradientVariant', AFillFormat.GradientVariant, RO);
      DisplayValue('Pattern', MsoPatternTypeStr(AFillFormat.Pattern), RO);
      Displayvalue('PresetGradientType', MsoPresetGradientTypeStr(AFillFormat.PresetGradientType), RO);
      DisplayValue('PresetTexture', MsoPresetTextureStr(AFillFormat.PresetTexture), RO);

      try
        DisplayValue('TextureName', AFillFormat.TextureName, RO);
      except
        on E:Exception do
          DisplayError('TextureName', E.Message);
      end;

      DisplayValue('TextureType', MsoTextureTypeStr(AFillFormat.TextureType), RO);
      DisplayValue('Transparency', AFillFormat.Transparency, RW);
{      Displayvalue('Type', MsoFillTypeStr(AFillFormat.Type), RO);}
      DisplayValue('Visible', MsoTriStateStr(AFillFormat.Visible), RO);
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayGroupItems(AGroupItems: OLEVariant);
begin
  if VarIsEmpty(AGroupItems) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      DisplayCreator(AGroupItems.Creator);
      DisplayValue('Count', AGroupItems.Count, RO);
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayHeaderFooter(AHeaderFooter: OLEVariant);
begin
  if VarIsEmpty(AHeaderFooter) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;

      try
        DisplayValue('Format', PpDateTimeFormatStr(AHeaderFooter.Format), RW);
      except
        on E:Exception do
          DisplayError('Format', E.Message);
      end;

      try
        DisplayValue('Text', AHeaderFooter.Text, RW);
      except
        on E:Exception do
          DisplayError('Text', E.Message);
      end;

      try
        DisplayValue('UseFormat', MsoTriStateStr(AHeaderFooter.UseFormat), RW);
      except
        on E:Exception do
          DisplayError('UseFormat', E.Message);
      end;

      try
        DisplayValue('Visible', MsoTriStateStr(AHeaderFooter.Visible), RW);
      except
        on E:Exception do
          DisplayError('Visible', E.Message);
      end;

    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayHeadersFooters(AHeadersFooters: OLEVariant);
begin
  if VarIsEmpty(AHeadersFooters) then
    exit;

  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      try
        DisplayValue('DisplayOnTitleSlide', MsoTriStateStr(AHeadersFooters.DisplayOnTitleSlide), RW);
      except
        on E:Exception do DisplayError('DisplayOnTitleSlide', E.Message);
      end;
    finally
      Items.Endupdate;
    end;
  end;
end;

procedure TMainForm.DisplayHyperlinks(AHyperlinks: OLEVariant);
var
  i: integer;
  s: string;
  h: OLEVariant;
begin
  DisplayValue('HyperLinks (Count)', AHyperLinks.Count, RO);
  s := '';
  for i := 1 to AHyperLinks.Count do
  begin
    h := AHyperLinks[i];
    if s = '' then
      s := h.Address
    else
      s := s + LineEnding + h.Address;
  end;
  DisplayValue('HyperLinks (Lines)', s, RO);
end;

procedure TMainForm.DisplayLineFormat(AFormat: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(AFormat) then
      begin
        DisplayCreator(AFormat.Creator);
        DisplayColor('BackColor', AFormat.BackColor, RW);
        DisplayValue('BeginArrowHeadLength', MsoArrowHeadLengthStr(AFormat.BeginArrowHeadLength), RW);
        DisplayValue('BeginArrowHeadStyle', MsoArrowHeadStyleStr(AFormat.BeginArrowHeadStyle), RW);
        DisplayValue('BeginArrowHeadWidth', MsoArrowHeadWidthStr(AFormat.BeginArrowHeadWidth), RW);
        DisplayValue('DashStyle', MsoLineDashStyleStr(AFormat.DashStyle), RW);
        DisplayValue('EndArrowHeadLength', MsoArrowHeadLengthStr(AFormat.EndArrowHeadLength), RW);
        DisplayValue('EndArrowHeadStyle', MsoArrowHeadStyleStr(AFormat.ENdArrowHeadStyle), RW);
        DisplayValue('EndArrowHeadWidth', MsoArrowHeadWidthStr(AFormat.EndArrowHeadWidth), RW);
        DisplayColor('ForeColor', AFormat.ForeColor, RW);
        DisplayValue('Pattern', MsoPatternTypeStr(AFormat.Pattern), RW);
        DisplayValue('Style', MsoLineStyleStr(AFormat.Style), RW);
        DisplayValue('Transparency', AFormat.Transparency, RW);
        DisplayValue('Visible', MsoTriStateStr(AFormat.Visible), RW);
        DisplayValue('Weight', AFormat.Weight, RW);
        DisplayValue('InsetPen', AFormat.InsetPen, RW);
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayLinkFormat(AFormat: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(AFormat) then
      begin
        try
          DisplayValue('AutoUpdate', PpUpdateOptionStr(AFormat.AutoUpdate), RW);
        except
          on E:Exception do
            DisplayError('AutoUpdate', E.Message);
        end;

        try
          DisplayValue('SourceFullName', AFormat.SourceFullName, RW);
        except
          on E:Exception do
            DisplayError('SourceFullName', E.Message);
        end;
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayMaster(AMaster: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(AMaster) then
      begin
        DisplayValue('Background', '*** ShapeRange not yet implemented ***', RO);
        DisplayValue('ColorScheme', '*** ColorScheme not yet implemented ***', RO);
        DisplayPoints('Height', AMaster.Height, RO);
        try
          DisplayHyperlinks(AMaster.HyperLinks);
        except
          // does not work for all master types
        end;
        DisplayValue('Name', AMaster.Name, RW);
        DisplayPoints('Width', AMaster.Width, RO);
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayMotionEffect(const Prefix: string; AEffect: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.ByX', AEffect.ByX, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ByX', E.Message);
  end;

  try
    DisplayValue(Prefix+'.ByY', AEffect.ByY, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ByY', E.Message);
  end;

  try
    DisplayValue(Prefix+'.FromX', AEffect.FromX, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.FromX', E.Message);
  end;

  try
    DisplayValue(Prefix+'.FromY', AEffect.FromY, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.FromY', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Path', AEffect.Path, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Path', E.Message);
  end;

  try
    DisplayValue(Prefix+'.ToX', AEffect.ToX, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ToX', E.Message);
  end;

  try
    DisplayValue(Prefix+'.ToY', AEffect.ToY, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ToY', E.Message);
  end;
end;

procedure TMainForm.DisplayOLEFormat(AOLEFormat: OLEVariant);
var
  s: string;
  i: integer;
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(AOLEFormat) then
      begin
        try
          DisplayValue('FollowColors', ppFollowColorsStr(AOLEFormat.FollowColors), RW);
        except
          on E:Exception do
            DisplayError('FollowColors', E.Message);
        end;

        try
          DisplayValue('ObjectVerbs (Count)', IntToStr(AOLEFormat.ObjectVerbs.Count), RO);
        except
          on E:Exception do
            DisplayError('ObjectVerbs.Count', E.Message);
        end;

        try
          s := '';
          for i:=0 to AOLEFormat.ObjectVers.Count do
          begin
            if s = '' then
              s := AOLEFormat.ObjectVerbs(i)
            else
              s := s + ' / ' + AOLEFormat.ObjectVerbs(i);
          end;
          DisplayValue('ObjectVerbs (Items)', s, RO);
        except
          on E:Exception do
            DisplayError('ObjectVerbs (Items)', E.Message);
        end;

        try
          DisplayValue('ProgID', AOLEFormat.ProgID, RO);
        except
          on E:Exception do
            DisplayError('ProgID', E.Message);
        end;
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayPageSetup(APageSetup: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(APageSetup) then
      begin
        DisplayValue('FirstSlideNumber', APageSetup.FirstSlideNumber, RW);
        DisplayValue('NotesOrientation', MsoOrientationStr(APageSetup.NotesOrientation), RW);
        DisplayPoints('SlideHeight', APageSetup.SlideHeight, RW);
        DisplayValue('SlideOrientation', MsoOrientationStr(APageSetup.SlideOrientation), RW);
        DisplayValue('SlideSize', PpSlideSizeTypeStr(APageSetup.SlideSize), RW);
        DisplayPoints('SlideWidth', APageSetup.SlideWidth, RW);
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayPictureFormat(APicFormat: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(APicFormat) then
      begin
        try
          DisplayValue('Brightness', APicFormat.Brightness, RW);
        except
          on E:Exception do
            DisplayError('Brightness', E.Message);
        end;

        try
          DisplayValue('ColorType', MsoPictureColorTypeStr(APicFormat.ColorType), RW);
        except
          on E:Exception do
            DisplayError('ColorType', E.Message);
        end;

        try
          Displayvalue('Contrast', APicformat.Contrast, RW);
        except
          on E:Exception do
            DisplayError('Contrast', E.Message);
        end;

        try
          DisplayCreator(APicFormat.Creator);
        except
        end;

        try
          DisplayPoints('CropBottom', APicFormat.CropBottom, RW);
        except
          on E:Exception do
            DisplayError('CropBottom', E.Message);
        end;

        try
          DisplayPoints('CropLeft', APicFormat.CropLeft, RW);
        except
          on E:Exception do
            DisplayError('CropLeft', E.Message);
        end;

        try
          DisplayPoints('CropRight', APicFormat.CropRight, RW);
        except
          on E:Exception do
            DisplayError('CropRight', E.Message);
        end;

        try
          DisplayPoints('CropTop', APicFormat.CropTop, RW);
        except
          on E:Exception do
            DisplayError('CropTop', E.Message);
        end;

        try
          DisplayValue('TransparencyColor', GetRGBName(APicFormat.TransparencyColor), RW);
        except
          on E:Exception do
            DisplayError('TransparencyColor', E.Message);
        end;

        try
          DisplayValue('TransparentBackground', MsoTriStateStr(APicFormat.TransparentBackground), RW);
        except
          on E:Exception do
            DisplayError('TransparentBackground', E.Message);
        end;
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayPlaySettings(const Prefix: string;
  ASettings: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.ActionVerb', ASettings.ActionVerb, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ActionVerb', E.Message);
  end;

  try
    DisplayValue(Prefix+'.HideWhileNotPlaying', MsoTriStateStr(ASettings.HideWhileNotPlaying), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.HideWhileNotPlaying', E.Message);
  end;

  try
    DisplayValue(Prefix+'.LoopUntilStopped', MsoTriStateStr(ASettings.LoopUntilStopped), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.LoopUntilStopped', E.Message);
  end;

  try
    DisplayValue(Prefix+'.PauseAnimation', MsoTriStateStr(ASettings.PauseAnimation), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.PauseAnimation', E.Message);
  end;

  try
    DisplayValue(Prefix+'.PlayOnEntry', MsoTriStateStr(ASettings.PlayOnEntry), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.PlayOnEntry', E.Message);
  end;

  try
    DisplayValue(Prefix+'.RewindMovie', MsoTriStateStr(ASettings.RewindMovie), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.RewindMovie', E.Message);
  end;

  try
    DisplayValue(Prefix+'.StopAfterSlides', ASettings.StopAfterSlides, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.StopAfterSlides', E.Message);
  end;
end;

procedure TMainForm.DisplayPoints(const ACaption: string; APoints: OLEVariant;
  AReadOnly: integer);
var
  P: double;
begin
  P := APoints;
  DisplayValue(ACaption, Format('%0.2f pts / %0.2f cm / %d pixels',
    [P, PointsToCm(P), PointsToPixels(P)]), AReadOnly);
end;

procedure TMainForm.DisplayPresentation(APresentation: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if VarIsEmpty(APresentation) then
        exit;

      try
        DisplayValue('DefaultLanguageID', MsoAppLanguageIDStr(APresentation.DefaultLanguageID), RW);
      except
        on E:Exception do DisplayError('DefaultLanguageID', E.Message);
      end;

      DisplayValue('DisplayComments', MsoTriStateStr(APresentation.DisplayComments), RW);
      try
        Displayvalue('EnvelopeVisible', APresentation.EnvelopeVisible, RW);
      except
        on E:Exception do DisplayError('EnvelopeVisible', E.Message);
      end;

      DisplayValue('ExtraColors', '*** ExtraColors noch nicht implementiert. ***', RO);
      Displayvalue('FullName', APresentation.FullName, RO);

      try
        DisplayPoints('GridDistance', APresentation.GridDistance, RW);
      except
        on E:Exception do DisplayError('GridDistance', E.Message);
      end;

      try
        DisplayValue('HasRevisionInfo', APresentation.HasRevisionInfo, RO);
      except
        on E:Exception do DisplayError('HasRevisionInfo', E.Message);
      end;

      DisplayValue('HasTitleMaster', MsoTriStateStr(APresentation.HasTitleMaster), RO);
      DisplayValue('LayoutDirection', APresentation.LayoutDirection, RW);
      DisplayValue('NoLineBreakAfter', APresentation.NoLineBreakAfter, RW);
      DisplayValue('NoLineBreakBefore', APresentation.NoLineBreakBefore, RW);

      try
        DisplayValue('Password', APresentation.Password, RW);
      except
        on E:Exception do DisplayError('Password', E.Message);
      end;

      try
        DisplayValue('PasswordEncryptionAlgorithm', APresentation.PasswordEncryptionAlgorithm, RO);
      except
        on E:Exception do DisplayError('PasswordEncryptionAlgorithm', E.Message);
      end;

      try
        Displayvalue('PasswordEncryptionKeyLength', APresentation.PasswordEncryptionKeyLength, RO);
      except
        on E:Exception do DisplayError('PasswordEncryptionKeyLength', E.Message);
      end;

      DisplayValue('Path', APresentation.Path, RO);
      DisplayValue('ReadOnly', MsoTriStateStr(APresentation.ReadOnly), RO);

      try
        DisplayValue('RemovePersonalInformation', APresentation.RemovePersonalInformation, RW);
      except
        on E:Exception do DisplayError('RemovePersonalInformation', E.Message);
      end;

      DisplayValue('Saved', MsoTriStateStr(APresentation.Saved), RW);
      try
        DisplayValue('SnapToGrid', MsoTriStateStr(APresentation.SnapToGrid), RW);
      except
        on E:Exception do DisplayError('SnapToGrid', E.Message);
      end;

      DisplayValue('TemplateName', APresentation.TemplateName, RO);
      try
        DisplayValue('VBASigned', APresentation.VBASigned, RO);
      except
        on E:Exception do DisplayError('VBASigned', E.Message);
      end;

      try
        DisplayValue('WritePassword', APresentation.WritePassword, RW);
      except
        on E:Exception do DisplayError('WritePassword', E.Message);
      end;

    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayPropertyEffect(const Prefix: string;
  AEffect: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.From', AEffect.From, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.From', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Points.Count', AEffect.Points.Count, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Points.Count', E.Message);
  end;

  try
    {DisplayValue(Prefix+'.Property', MsoAnimPropertyStr(AEffect.Property), RW);}
  except
    on E:Exception do
      DisplayError(Prefix+'.Property', E.Message);
  end;

  try
    DisplayValue(Prefix+'.To', AEffect.&To, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.To', E.Message);
  end;
end;

procedure TMainForm.DisplayRGB(const Prefix: string; RGB: OLEVariant;
  AReadOnly: integer);
var
  n: integer;
begin
  n := RGB;
  DisplayValue(Prefix + '.RGB',
    Format('$%.8x (%s)', [n, GetRGBName(n)]), AReadOnly, 'C');
end;

procedure TMainForm.DisplayRotationEffect(const Prefix: string;
  AEffect: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.By', AEffect.By, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.By', E.Message);
  end;

  try
    DisplayValue(Prefix+'.From', AEffect.From, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.From', E.Message);
  end;

  try
    DisplayValue(Prefix+'.To', AEffect.&To, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.To', E.Message);
  end;
end;

procedure TMainForm.DisplayScaleEffect(const Prefix: string;
  AEffect: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.ByX', AEffect.ByX, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ByX', E.Message);
  end;

  try
    DisplayValue(Prefix+'.ByY', AEffect.ByY, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ByY', E.Message);
  end;

  try
    DisplayValue(Prefix+'.FromX', AEffect.FromX, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.FromX', E.Message);
  end;

  try
    DisplayValue(Prefix+'.FromY', AEffect.FromY, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.FromY', E.Message);
  end;

  try
    DisplayValue(Prefix+'.ToX', AEffect.ToX, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ToX', E.Message);
  end;

  try
    DisplayValue(Prefix+'.ToY', AEffect.ToY, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.ToY', E.Message);
  end;
end;

// Only for Office 2002+
procedure TMainForm.DisplaySequence(ASequence: OLEVariant);
begin
  with ValueList do
  begin
    if not VarIsEmpty(ASequence) then
    begin
      try
        DisplayValue('Count', ASequence.Count, RO);
      except
        on E:Exception do
          DisplayError('Count', E.Message);
      end;
    end;
  end;
end;

procedure TMainForm.DisplayShadow(AShadow: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not varIsEmpty(AShadow) then
      begin
        DisplayCreator(AShadow.Creator);
        DisplayColor('ForeColor', AShadow.ForeColor, RW);
        DisplayValue('Obscured', MsoTriStateStr(AShadow.Obscured), RW);
        DisplayValue('OffsetX', AShadow.OffsetX, RW);
        DisplayValue('OffsetY', AShadow.OffsetY, RW);
        DisplayValue('Transparency', AShadow.Transparency, RW);
{        DisplayValue('Type', MsoShadowTypeStr(AShadow.Type), RW);}
        DisplayValue('Visible', MsoTriStateStr(AShadow.Visible), RW);
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayShape(AShape: OLEVariant);
var
  n, w, h : integer;
begin
  with ValueList do begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(AShape) then begin
        try
          DisplayValue('AlternativeText', AShape.AlternativeText, RW);
        except
          on E:Exception do DisplayError('AlternativeText', E.Message);
        end;

        try
          DisplayValue('AutoShapeType', MsoAutoShapeTypeStr(AShape.AutoShapeType), RW);
        except
          on E:Exception do DisplayError('AutoShapeType', E.Message);
        end;

        try
          DisplayValue('BlackWhiteMode', MsoBlackWhiteModeStr(AShape.BlackWhiteMode), RW);
        except
          on E:Exception do DisplayError('BlackWhiteMode', E.Message);
        end;

        try
          DisplayValue('Child', MsoTriStateStr(AShape.Child), RO);
        except
          on E:Exception do DisplayError('Child', E.Message);
        end;

        try
          DisplayValue('ConnectionSiteCount', AShape.ConnectionSiteCount, RO);
        except
          on E:Exception do DisplayError('ConnectionSiteCount', E.Message);
        end;

        try
          DisplayValue('Connector', MsoTriStateStr(AShape.Connector), RW);
        except
          on E:Exception do DisplayError('Connector', E.Message);
        end;

        try
          DisplayCreator(AShape.Creator);
        except
        end;

        try
          DisplayValue('HasDiagram', MsoTriStateStr(AShape.HasDiagram), RO);
        except
          on E:Exception do DisplayError('HasDiagram', E.Message);
        end;

        try
          DisplayValue('HasDiagramNode', MsoTriStateStr(AShape.HasDiagramNode), RO);
        except
          on E:Exception do DisplayError('HasDiagramNode', E.Message);
        end;

        try
          DisplayValue('HasTable', MsoTriStateStr(AShape.HasTable), RO);
        except
          on E:Exception do DisplayError('HasTable', E.Message);
        end;

        try
          DisplayValue('HasTextFrame', MsoTriStateStr(AShape.HasTextFrame), RO);
        except
          on E:Exception do DisplayError('HasTextFrame', E.Message);
        end;

        try
          DisplayPoints('Height', AShape.Height, RW);
        except
          on E:Exception do DisplayError('Height', E.Message);
        end;

        try
          DisplayValue('HorizontalFlip', MsoTriStateStr(AShape.HorizontalFlip), RO);
        except
          on E:Exception do DisplayError('HorizontalFlip', E.Message);
        end;

        try
          n := AShape.ID;
          DisplayValue('ID', Format('$%x (%d)', [n, n]), RO);
        except
          on E:Exception do DisplayError('ID', E.Message);
        end;

        try
          DisplayPoints('Left', AShape.Left, RW);
        except
          on E:Exception do DisplayError('Left', E.Message);
        end;

        try
          DisplayValue('LockAspectRatio', MsoTriStateStr(AShape.LockAspectRatio), RW);
        except
          on E:Exception do DisplayError('LockAspectRatio', E.Message);
        end;

        try
          DisplayValue('MediaType', PpMediaTypeStr(AShape.MediaType), RO);
        except
          on E:Exception do DisplayError('MediaType', E.Message);
        end;

        try
          DisplayValue('Name', AShape.Name, RW);
        except
          on E:Exception do DisplayError('Name', E.Message);
        end;

        try
          DisplayValue('PlaceholderFormat.Type', PpPlaceHolderTypeStr(AShape.PlaceHolderFormat.&Type), RO);
        except
          on E:Exception do DisplayError('PlaceholderFormat.Type', E.Message);
        end;

        try
          DisplayValue('Rotation', AShape.Rotation, RW);
        except
          on E:Exception do DisplayError('Rotation', E.Message);
        end;

        try
          DisplayPoints('Top', AShape.Top, RW);
        except
          on E:Exception do DisplayError('Top', E.Message);
        end;

        try
          DisplayValue('Type', MsoShapeTypeStr(AShape.&Type), RO);
        except
          on E:Exception do DisplayError('Type', E.Message);
        end;

        try
          DisplayValue('VerticalFlip', MsoTriStateStr(AShape.VerticalFlip), RO);
        except
          on E:Exception do DisplayError('VerticalFlip', E.Message);
        end;

        try
          DisplayPoints('Width', AShape.Width, RW);
        except
          on E:Exception do DisplayError('Width', E.Message);
        end;

        try
          DisplayValue('ZOrderPosition', AShape.ZOrderPosition, RO);
        except
          on E:Exception do DisplayError('ZOrderPosition', E.Message);
        end;

        if AShape.&Type=msoPicture then begin
          OriginalPictureSize(AShape, w, h);
          DisplayValue('*** Original Picture Size ***', Format('%d x %d picels', [w, h]), RO);
        end;

  //      DisplayProperty(' *** MemoryUsage ***', Format('%0.2n KB', [1.0*GetShapeMemSize(AShape)/1024]));
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplayShapes(AShapes:OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(AShapes) then
      begin
        DisplayCreator(AShapes.Creator);
        DisplayValue('Count', AShapes.Count, RO);
        DisplayValue('HasTitle', MsoTriStateStr(AShapes.hasTitle), RO);
        try
          DisplayTextRange('Title.TextFrame', AShapes.Title.TextFrame.TextRange);
        except
          on E:Exception do
            DisplayError('Title.TextFrame', E.Message);
        end;
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplaySlide(ASlide: OLEVariant);
var
  s: string;
  i: integer;
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(ASlide) then
      begin
        try
          DisplayValue('Comments (Count)', ASlide.Comments.Count, RO);
        except
          on E:Exception do DisplayError('Comments (Count)', E.Message);
        end;

        try
          s := '';
          for i:=1 to ASlide.Comments.Count do
            if s='' then
              s := ASlide.Comments[i]
            else
              s := s + LineEnding + ASlide.Comments[i];
          DisplayValue('Comments (Lines)', s, RO);
        except
          on E:Exception do DisplayError('Comments (Lines)', E.Message);
        end;

        DisplayValue('DisplayMasterShapes', MsoTriStateStr(ASlide.DisplayMasterShapes), RW);
        DisplayValue('FollowMasterBackground', MsoTriStateStr(ASlide.FollowMasterBackground), RW);
        DisplayHyperlinks(ASlide.Hyperlinks);
        DisplayValue('Layout', PpSlideLayoutStr(ASlide.Layout), RW);
        DisplayValue('Master.Name', ASlide.Master.Name, RO);
        DisplayValue('Name', ASlide.Name, RW);
        DisplayValue('PrintSteps', ASlide.PrintSteps, RO);
        DisplayValue('SlideID', ASlide.SlideID, RO);
        DisplayValue('SlideNumber', ASlide.SlideNumber, RO);
        DisplayValue('SlideIndex', ASlide.SlideIndex, RO);
  //      DisplayValue(' *** MemoryUsage ***', Format('%0.2n KB', [1.0*GetSlideMemSize(ASlide)/1024]));
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplaySlides(ASlides: OLEVariant);
begin
  with ValueList do
  begin
    Items.BeginUpdate;
    try
      Items.Clear;
      if not VarIsEmpty(ASlides) then
      begin
        DisplayValue('Count', ASlides.Count, RO);
      end;
    finally
      Items.EndUpdate;
    end;
  end;
end;

procedure TMainForm.DisplaySlideShowSettings(ASettings: OLEVariant);
begin
  if VarIsEmpty(ASettings) then
    exit;

  with ValueList do
  begin
    DisplayValue('AdvanceMode', PpAdvanceModeStr(ASettings.AdvanceMode), RW);
    DisplayValue('EndingSlide', ASettings.EndingSlide, RW);
    DisplayValue('LoopUntilStopped', MsoTriStateStr(ASettings.LoopUntilStopped), RW);
    try
      DisplayColor('PointerColor', ASettings.PointerColor, RO);
    except
      on E:Exception do DisplayError('PointerColor', E.Message);
    end;
    DisplayValue('RangeType', PpSlideshowRangeTypeStr(ASettings.RangeType), RW);
    try
      DisplayValue('ShowScrollbar', MsoTriStateStr(ASettings.ShowScrollbar), RW);
    except
      on E:Exception do DisplayError('ShowScrollbar', E.Message);
    end;
    DisplayValue('ShowType', PpSlideShowTypeStr(ASettings.ShowType), RW);
    DisplayValue('ShowWithNarration', MsoTriStateStr(ASettings.ShowWithNarration), RW);
    DisplayValue('ShowWithAnimation', MsoTriStateStr(ASettings.ShowWithAnimation), RW);
    DisplayValue('SlideshowName', ASettings.SlideShowName, RO);
    DisplayValue('StartingSlide', ASettings.StartingSlide, RW);
  end;
end;

procedure TMainForm.DisplaySlideshowTransition(ASettings: OLEVariant);
begin
  if VarIsEmpty(ASettings) then
    exit;

  with ValueList do
  begin
    DisplayValue('AdvanceOnClick', MsoTristateStr(ASettings.AdvanceOnClick), RW);
    DisplayValue('AdvanceOnTime', MsoTriStateStr(ASettings.AdvanceOnTime), RW);
    DisplayValue('AdvanceTime', ASettings.AdvanceTime, RW);
    DisplayValue('EntryEffect', PpEntryEffectStr(ASettings.EntryEffect), RW);
    DisplayValue('Hidden', MsoTriStateStr(ASettings.Hidden), RW);
    DisplayValue('LoopSoundUntilNext', MsoTriStateStr(ASettings.LoopSoundUntilNext), RW);
    DisplaySoundEffect(ASettings.SoundEffect);
    DisplayValue('Speed', ASettings.Speed, RW);
  end;
end;

procedure TMainForm.DisplaySoundEffect(AEffect:OLEVariant);
begin
  try
    DisplayValue('SoundEffect.Name', PpSoundEffectTypeStr(AEffect.Name), RW);
  except
    DisplayError('SoundEffect.Name', '*** kein Sound ***');
  end;
  DisplayValue('SoundEffect.Type', PpSoundEffectTypeStr(AEffect.&Type), RW);
end;

procedure TMainForm.DisplayTextEffect(ATextEffect: OLEVariant);
begin
  with ValueList do
  begin
    if not VarIsEmpty(ATextEffect) then
    begin
      try
        Displayvalue('Alignment', MsoTextEffectAlignmentStr(ATextEffect.Alignment), RW);
      except
        on E:Exception do DisplayError('Alignment', E.Message);
      end;

      try
        DisplayCreator(ATextEffect.Creator);
      except
        on E:Exception do DisplayError('Creator', E.Message);
      end;

      try
        Displayvalue('FontBold', MsoTriStateStr(ATextEffect.FontBold), RW);
      except
        on E:Exception do DisplayError('FontBold', E.Message);
      end;

      try
        DisplayValue('FontItalic', MsoTriStateStr(ATextEffect.FontItalic), RW);
      except
        on E:Exception do DisplayError('FontItalic', E.Message);
      end;

      try
        DisplayValue('FontName', ATextEffect.FontName, RW);
      except
        on E:Exception do DisplayError('FontName', E.Message);
      end;

      try
        DisplayValue('FontSize', ATextEffect.FontSize, RW);
      except
        on E:Exception do DisplayError('FontSize', E.Message);
      end;

      try
        DisplayValue('KernedPairs', MsoTriStateStr(ATextEffect.KernedPairs), RW);
      except
        on E:Exception do DisplayError('KernedPairs', E.Message);
      end;

      try
        DisplayValue('NormalizedHeight', MsoTriStateStr(ATextEffect.NormalizedHeight), RW);
      except
        on E:Exception do DisplayError('NormalizedHeight', E.Message);
      end;

      try
        DisplayValue('RotatedChars', MsoTriStateStr(ATextEffect.RotatedChars), RW);
      except
        on E:Exception do DisplayError('RotatedChars', E.Message);
      end;

      try
        DisplayValue('Text', ATextEffect.Text, RW);
      except
        on E:Exception do DisplayError('Text', E.Message);
      end;

      try
        DisplayValue('Tracking', ATextEffect.Tracking, RW);
      except
        on E:Exception do DisplayError('Tracking', E.Message);
      end;
    end;
  end;
end;

procedure TMainForm.DisplayTextFrame(ATextFrame: OLEVariant);
begin
  with ValueList do
  begin
    if not VarIsEmpty(ATextFrame) then
    begin
      try
        DisplayValue('AutoSize', ppAutoSizeStr(ATextFrame.AutoSize), RW);
      except
        on E:Exception do DisplayError('AutoSize', E.Message);
      end;

      try
        DisplayCreator(ATextFrame.Creator);
      except
      end;

      try
        DisplayValue('HasText', MsoTriStateStr(ATextFrame.HasText), RO);
      except
        on E:Exception do DisplayError('HasText', E.Message);
      end;

      try
        DisplayValue('HorizontalAnchor', ATextFrame.HorizontalAnchor, RW);
      except
        on E:Exception do DisplayError('HorizontalAnchor', E.Message);
      end;

      try
        DisplayPoints('MarginBottom', ATextFrame.MarginBottom, RW);
      except
        on E:Exception do DisplayError('MarginBottom', E.Message);
      end;

      try
        DisplayPoints('MarginLeft', ATextFrame.MarginLeft, RW);
      except
        on E:Exception do DisplayError('MarginLeft', E.Message);
      end;

      try
        DisplayPoints('MarginRight', ATextFrame.MarginRight, RW);
      except
        on E:Exception do DisplayError('MarginRight', E.Message);
      end;

      try
        DisplayPoints('MarginTop', ATextFrame.MarginTop, RW);
      except
        on E:Exception do DisplayError('MarginTop', E.Message);
      end;

      try
        DisplayValue('Orientation', MsoTextOrientationStr(ATextFrame.Orientation), RW);
      except
        on E:Exception do DisplayError('Orientation', E.Message);
      end;

      try
        DisplayTextRange('', ATextFrame.TextRange);
      except
      end;

      try
        DisplayValue('VerticalAnchor', ATextFrame.VerticalAnchor, RW);
      except
        on E:Exception do DisplayError('VerticalAnchor', E.Message);
      end;

      try
        Displayvalue('WordWrap', MsoTriStateStr(ATextFrame.WordWrap), RW);
      except
        on E:Exception do DisplayError('WordWrap', E.Message);
      end;
    end;
  end;
end;

procedure TMainForm.DisplayTextRange(const PreFix: string;
  ATextRange: OLEVariant);
begin
  with ValueList do
  begin
    DisplayPoints(Prefix+'TextRange.BoundHeight', ATextRange.BoundHeight, RO);
    DisplayPoints(Prefix+'TextRange.BoundLeft', ATextRange.BoundLeft, RO);
    DisplayPoints(Prefix+'TextRange.BoundTop', ATextRange.BoundTop, RO);
    DisplayPoints(Prefix+'TextRange.BoundWidth', ATextRange.BoundWidth, RO);
    Displayvalue(Prefix+'TextRange.Count', ATextRange.Count, RO);
    Displayvalue(Prefix+'TextRange.IndentLevel', ATextRange.IndentLevel, RW);
    try
      DisplayValue(Prefix+'TextRange.LanguageID', MsoAppLanguageIDStr(ATextRange.LanguageID), RW);
    except
      on E:Exception do DisplayError(Prefix+'.TextRange.LanguageID', E.Message);
    end;
    DisplayValue(Prefix+'TextRange.Length', ATextRange.Length, RO);
    DisplayValue(Prefix+'TextRange.Start', ATextRange.Start, RO);
    DisplayValue(Prefix+'TextRange.Text', ATextRange.Text, RW);
  end;
end;

procedure TMainForm.DisplayTiming(const Prefix: string; AValue: OLEVariant);
begin
  try
    DisplayValue(Prefix+'.AccelerateX', AValue.Accelerate, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Acceleratore', E.Message);
  end;

  try
    DisplayValue(Prefix+'.AutoReverse', MsoTriStateStr(AValue.AutoReverse), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.AutoReverse', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Decelerate', AValue.Decelerate, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Decelerate', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Duration', AValue.Duration, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Duration', E.Message);
  end;

  try
    DisplayValue(Prefix+'.RepeatCount', AValue.RepeatCount, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.RepeatCount', E.Message);
  end;

  try
    DisplayValue(Prefix+'.RepeatDuration', AValue.RepeatDuration, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.RepeatDuration', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Restart', MsoAnimEffectRestartStr(AValue.Restart), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Restart', E.Message);
  end;

  try
    DisplayValue(Prefix+'.RewindAtEnd', MsoTriStateStr(AValue.RewindAtEnd), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.RewindAtEnd', E.Message);
  end;

  try
    DisplayValue(Prefix+'.SmoothEnd', MsoTriStateStr(AValue.SmoothEnd), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.SmoothEnd', E.Message);
  end;

  try
    DisplayValue(Prefix+'.SmoothStart', MsoTriStateStr(AValue.SmoothStart), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.SmoothStart', E.Message);
  end;

  try
    DisplayValue(Prefix+'.Speed', AValue.Speed, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.Speed', E.Message);
  end;

  try
    DisplayValue(Prefix+'.TriggerDelayTime', AValue.TriggerDelayTime, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.TriggerDelayTime', E.Message);
  end;

  try
    DisplayValue(Prefix+'.TriggerShape.Name', AValue.TriggerShape.Name, RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.TriggerShape.Name', E.Message);
  end;

  try
    DisplayValue(Prefix+'.TriggerType', MsoAnimTriggerTypeStr(AValue.TriggerType), RW);
  except
    on E:Exception do
      DisplayError(Prefix+'.TriggerType', E.Message);
  end;
end;

procedure TMainForm.DisplayValue(const AName: string; AValue: variant;
  AReadOnly: integer; ADisplayOption: char = #0; AEditOption: char = #0);
var
  Item: TListItem;
  s: string;
begin
  Item := ValueList.Items.Add;
  with Item do begin
    Caption := AName;
    s := AValue;
    SubItems.Add(s);
    SubItems.Add(ADisplayOption + AEditOption);
    if (Length(s) > 1) and (s[1] = '-') and (s[2] = '-') then
      ImageIndex := IMG_INDEX_ERROR       //26
    else if AReadOnly = RO then
      ImageIndex := IMG_INDEX_READONLY    //22
    else
      ImageIndex := IMG_INDEX_READWRITE;  // 21
  end;
end;

function TMainForm.FindPresentation(ANode: TTreeNode): OleVariant;
var
  idx: Integer;
begin
  ANode := FindPresentationNode(ANode);
  idx := abs(Integer(PtrInt(ANode.Data)));
  Result := FPowerpoint.Presentations.Item(idx);
end;

function TMainForm.FindPresentationNode(ANode: TTreeNode): TTreeNode;
var
  N: TTreeNode;
begin
  N := ANode;
  while Assigned(N) do
  begin
    N := ANode.Parent;
    if N = nil then
    begin
      Result := ANode;
      exit;
    end;
    ANode := N;
  end;
  Result := nil;
end;

function TMainForm.FindShape(ANode: TTreeNode): OLEVariant;
var
  slide: OLEVariant;
  slideID: integer;
  slideNode: TTreeNode;
  masterNode: TTreeNode;
  N: TTreeNode;
begin
  Result := Unassigned;
  if Assigned(ANode) then
  begin
    if (ANode.Level = 4) then
    begin
      slideNode := nil;
      masterNode := nil;
      if (ANode.Parent.Parent.Parent.Text = 'Slides') then
        slideNode := ANode.Parent.Parent
      else if (ANode.Parent.Parent.Text = 'TitleMaster')
        or (ANode.Parent.Parent.Text = 'NotesMaster')
        or (ANode.Parent.Parent.Text = 'SlideMaster')
        or (ANode.Parent.Parent.Text = 'HandoutMaster')
      then
        masterNode := ANode.Parent.Parent;
    end else
    if (ANode.Level = 6) then
    begin
      N := ANode.Parent.Parent.Parent.Parent;
      if (N.Parent.Text = 'Slides') then
        slideNode := N
      else if (N.Text = 'TitleMaster') or (N.Text = 'NotesMaster')
        or (N.Text = 'SlideMaster') or (N.Text = 'HandoutMaster')
      then
        masterNode := N;
    end;

    if Assigned(slideNode) then
    begin
      slideID := integer(slideNode.Data);
      slide := FActPresentation.Slides.FindBySlideID(slideID);
      Result := FindShape(slide.Shapes, ANode.Text);
    end else
    if Assigned(MasterNode) then
    begin
      if MasterNode.Text = 'TitleMaster' then
        Result := FindShape(FActPresentation.TitleMaster.Shapes, ANode.Text)
      else if (MasterNode.Text = 'HandoutMaster') then
        Result := FindShape(FActPresentation.HandoutMaster.Shapes, ANode.Text)
      else if (MasterNode.Text = 'NotesMaster') then
        Result := FindShape(FActPresentation.NotesMaster.Shapes, ANode.Text)
      else if (MasterNode.Text='SlideMaster') then
        Result := FindShape(FActPresentation.SlideMaster.Shapes, ANode.Text);
    end;
  end;
end;

function TMainForm.FindShape(AShapeCollection: OLEVariant;
  const AShapeName: string): OLEVariant;
var
  i : integer;
  shape : OLEVariant;
begin
  Result := UnAssigned;
  for i := 1 to AShapeCollection.Count do
  begin
     shape := AShapeCollection.Item(i);
     if CompareText(shape.Name, AShapeName) = 0 then
     begin
       Result := shape;
       exit;
     end else
     if (shape.&Type=msoGroup) then begin
       result := FindShape(shape.GroupItems, AShapeName);
       if not VarIsEmpty(result) then exit;
     end;
  end;
end;

function TMainForm.FindSlide(ANode: TTreeNode): OLEVariant;
begin
  Result := UnAssigned;
  while Assigned(ANode) and Assigned(ANode.Parent) do
  begin
    if ANode.Parent.Text = 'Slides' then
    begin
      Result := FActPresentation.Slides.FindBySlideID(PtrInt(ANode.Data));
      exit;
    end;
    ANode := ANode.Parent;
  end;
end;

function TMainForm.FindSlide(ASlideName: string): OLEVariant;
var
  crs: TCursor;
  i: integer;
  slides: OLEVariant;
  slide: OLEVariant;
begin
  Result := UnAssigned;
  crs := Screen.Cursor;
  Screen.Cursor := crHourglass;
  try
    i := pos('(#', ASlideName);
    if i > 0 then Delete(ASlideName, i-1, Length(ASlideName));
    slides := FActPresentation.Slides;
    for i := 1 to slides.Count do
    begin
      slide := slides.Item(i);
      if CompareText(ASlideName, slide.Name) = 0 then
      begin
        Result := slide;
        exit;
      end;
    end;
  finally
    Screen.Cursor := crs;
  end;
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  somedirty: boolean;
  N: TTreeNode;
begin
  somedirty := false;

  if Treeview.Items.Count > 0 then
  begin
    N := Treeview.Items[0];
    while (N <> nil) do
    begin
      if IsDirty(N) then
      begin
        somedirty := true;
        break;
      end;
      N := N.GetNextSibling;
    end;

    if somedirty then
    begin
      Treeview.Selected := N;
      case MessageDlg('This PowerPoint file has been modified. Save?', mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
        mrYes:
          begin
            AcFileSaveAs.Execute;
            CanClose := false;
          end;
        mrNo:
          CanClose := true;
        mrCancel:
          CanClose := false;
      end;
    end;
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
  R: TRect;
  bmp: TBitmap;
begin
  FActSlideID := -1;

  ReadIni;

  try
    FShowErrorItems := CbShowErrorItems.Checked;
    SystemParametersInfo(SPI_GETWORKAREA, 0, @R, 0);
//    R := Screen.WorkAreaRect;
    SetBounds(R.Left, R.Top, R.Bottom - R.Top, (R.Right - R.Left) div 2);
    try
      FPowerPoint := GetActiveOleObject('PowerPoint.Application');
      FPowerpointRunning := true;
    except
      try
        FPowerPoint := CreateOleObject('PowerPoint.Application');
        FPowerpointRunning := false;
      except
        FPowerpointRunning := false;
        raise Exception.Create('Powerpoint is not installed on this machine.');
      end;
    end;

    FPowerpoint.Visible := true;
    FPowerpointWidth := FPowerpoint.Width;
    FPowerpointHeight := FPowerpoint.Height;
    FPowerpointLeft := FPowerpoint.Left;
    FPowerpointTop := FPowerpoint.Top;
    try
      FPowerpoint.Left := PixelsToPoints(Width);
      FPowerpoint.Top := PixelsToPoints(R.Top);
      FPowerpoint.Width := PixelsToPoints(Screen.Width) - Powerpoint.Left;
      FPowerpoint.Height := PixelsToPoints(Height);
    except
      // Exception raised when PPT already has been opened maximized.
    end;

    FPowerpoint.Presentations.Add(msoFalse);
    FTestPresentation := FPowerpoint.Presentations.Item(FPowerpoint.Presentations.Count);

    bmp := TBitmap.Create;
    try
      TreeviewImages.GetBitmap(IMG_INDEX_READONLY, bmp);
      ImgReadOnly.Picture.Assign(bmp);
    finally
      bmp.Free;
    end;

    bmp := TBitmap.Create;
    try
      TreeViewImages.GetBitmap(IMG_INDEX_READWRITE, bmp);
      ImgReadWrite.Picture.Assign(bmp);
    finally
      bmp.Free;
    end;

  except
    raise Exception.Create('PowerPoint is not installed on this system.');
  end;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  WriteIni;

  try
    FTestPresentation.Close;
    FTestPresentation := UnAssigned;
  except
  end;

  try
    FPowerpoint.Left := FPowerPointLeft;
    FPowerpoint.Top := FPowerPointTop;
    FPowerpoint.Width := FPowerPointWidth;
    FPowerpoint.Height := FPowerPointHeight;
  except
    // Exception entsteht, wenn PPT maximiert ist.
  end;

  if not FPowerpointRunning then
    FPowerPoint.Quit;
  FPowerPoint := UnAssigned;
end;

procedure TMainForm.MRUMenuManagerRecentFile(Sender: TObject;
  const AFileName: String);
begin
  AddPresentation(AFileName);
end;

procedure TMainForm.TreeViewChange(Sender: TObject; Node: TTreeNode);
var
  shape: OLEVariant;
  slide: OLEVariant;
  master: OLEVariant;
  hdrFtr: OLEVariant;
  view: OLEVariant;
  seq: OLEVariant;
  eff: OLEVariant;
  crs: TCursor;
  s: string;
begin
  crs := Screen.Cursor;
  Screen.Cursor := crHourglass;
  ValueList.Items.BeginUpdate;
  try
    ValueList.Items.Clear;

    case Node.Level of
      0 : begin
            SelectPresentation(Node);
            DisplayPresentation(FActPresentation);
          end;

      1 : if Node.Text = 'Slides' then
          begin
            SelectPresentation(Node);
            DisplaySlides(FActPresentation.Slides);
          end else
          if Node.Text = 'BuiltInDocumentProperties' then
          begin
            SelectPresentation(Node);
            DisplayDocProperties(FActPresentation.BuiltInDocumentProperties);
          end else
          if Node.Text = 'CustomDocumentProperties' then
          begin
            SelectPresentation(Node);
            DisplayDocProperties(FActPresentation.CustomDocumentProperties);
          end else
          if Node.Text = 'DefaultShape' then
          begin
            SelectPresentation(Node);
            DisplayShape(FActPresentation.DefaultShape)
          end else
          if Node.Text = 'SlideshowSettings' then
          begin
            SelectPresentation(Node);
            DisplaySlideshowSettings(FActPresentation.SlideShowSettings);
          end else
          if Node.Text = 'PageSetup' then
          begin
            SelectPresentation(Node);
            DisplayPageSetup(FActPresentation.PageSetup);
          end;

      2 : if (Node.Parent.Text = 'Slides') then
          begin
            slide := SelectSlide(Node);
            DisplaySlide(slide);
          end else
          if (Node.Parent.Text = 'Masters') then
          begin
            master := SelectMaster(Node);
            DisplayMaster(master);
          end;

      3 : if (Node.Parent.Parent.Text = 'Slides') or
             (Node.Parent.Parent.Text = 'Masters') then
          begin
            if (Node.Text = 'HeadersFooters') then
            begin
              if (Node.Parent.Parent.Text = 'Slides') then
              begin
                slide := SelectSlide(Node.Parent);
                DisplayHeadersFooters(slide.HeadersFooters);
              end else
              if (Node.Parent.Parent.Text = 'Masters') then
              begin
                master := SelectMaster(Node.Parent);
                DisplayHeadersFooters(master.HeadersFooters);
              end;
            end else
            if (Node.Text = 'Shapes') then
            begin
              if (Node.Parent.Parent.Text = 'Slides') then
              begin
                slide := SelectSlide(Node.Parent);
                DisplayShapes(slide.Shapes);
              end else
              if (Node.Parent.Parent.Text = 'Masters') then
              begin
                master := SelectMaster(Node.Parent);
                DisplayShapes(master.Shapes);
              end;
            end else
            if (Node.Text = 'SlideshowTransition') then
            begin
              if Node.Parent.Parent.Text = 'Masters' then
              begin
                master := SelectMaster(Node.Parent);
                DisplaySlideshowTransition(master.Slideshowtransition);
              end else
              if (Node.Parent.Parent.Text = 'Slides') then
              begin
                slide := SelectSlide(Node.Parent);
                DisplaySlideShowTransition(slide.SlideShowTransition);
              end;
            end;
          end;

      4 : if (Node.Parent.Text = 'HeadersFooters') then
          begin
            if (Node.Parent.Parent.Parent.Text = 'Slides')
              or (Node.Parent.Parent.Parent.Text = 'Masters') then
            begin
              try
                hdrftr := SelectHeaderFooter(Node);
                DisplayHeaderFooter(hdrftr);
              except
                exit;
              end;
            end;
          end else
          if (Node.Parent.Text = 'Shapes') then
          begin
            shape := SelectShape(Node);
            DisplayShape(shape);
          end else
          if (Node.Parent.Text = 'Sequences') then
          begin
            seq := SelectSequence(Node);
            DisplaySequence(seq);
          end;

      5 : if (Node.Parent.Parent.Text = 'Shapes') then begin
            shape := SelectShape(Node.Parent);
            if (Node.Text = 'ActionSettings (ppMouseClick)') then begin
              if shape.&Type <> msoGroup then
                DisplayActionSettings(shape.ActionSettings, ppMouseClick);
            end else
            if (Node.Text = 'ActionSettings (ppMouseOver)') then begin
              if shape.&Type <> msoGroup then
                DisplayActionSettings(shape.ActionSettings, ppMouseOver);
            end else
            if (Node.Text = 'AnimationSettings') then
              DisplayAnimationSettings(shape.AnimationSettings)
            else
            if (Node.Text = 'Callout') then
              DisplayCallout(shape.Callout)
            else
            if (Node.Text = 'Diagram') then
            begin
              try
                if shape.HasDiagram then
                  DisplayDiagram(shape.Diagram)
              except
              end
            end else
            if (Node.Text = 'LinkFormat') then
              DisplayLinkFormat(shape.LinkFormat)
            else
            if (Node.Text = 'OLEFormat') then
              DisplayOLEFormat(shape.OLEFormat)
            else
            if (Node.Text = 'PictureFormat') then
              DisplayPictureFormat(shape.PictureFormat)
            else
            if (Node.Text = 'Shadow') then
              DisplayShadow(shape.Shadow)
            else
            if (Node.Text = 'Fill') then
              DisplayFillFormat(shape.Fill)
            else
            if (Node.Text = 'TextEffect') then
              DisplayTextEffect(shape.TextEffect)
            else
            if (Node.Text = 'TextFrame') then
              DisplayTextFrame(shape.TextFrame)
            else
            if (Node.Text = 'GroupItems') then
              DisplayGroupItems(shape.GroupItems);
          end else
          if Node.Parent.Parent.Text = 'Sequences' then
          begin
            eff := SelectEffect(Node);
            DisplayEffect(eff);
          end;

      6 : if (Node.Parent.Text = 'GroupItems') then
          begin
            shape := SelectShape(Node);
            DisplayShape(shape);
          end else
          if (Node.Parent.Parent.Parent.Text = 'Sequences') then
          begin
            eff := SelectEffect(Node.Parent);
            if Node.Text = 'EffectInformation' then
              DisplayEffectInformation(eff.EffectInformation)
            else
            begin
              s := copy(Node.Text, 1, pos('.', Node.Text));
              if s = 'Behaviors.' then
              begin
                s := copy(Node.Text, pos('(', Node.Text)+1, 255);
                Delete(s, Length(s), 1);
                eff := SelectEffect(Node.Parent);
                DisplayAnimationBehavior(eff.Behaviors.Item(StrToInt(s)));
              end;
            end;
          end;

      7 : if (Node.Parent.Parent.Text='GroupItems') then
          begin
            shape := SelectShape(Node.Parent);
            if (Node.Text = 'ActionSettings (ppMouseClick)') then
            begin
              if shape.&Type <> msoGroup then
                DisplayActionSettings(shape.ActionSettings, ppMouseClick);
            end else
            if (Node.Text = 'ActionSettings (ppMouseOver)') then
            begin
              if shape.&Type <> msoGroup then
                DisplayActionSettings(shape.ActionSettings, ppMouseOver);
            end else
            if (Node.Text = 'AnimationSettings') then
              DisplayAnimationSettings(shape.AnimationSettings)
            else
            if (Node.Text = 'Callout') then
              DisplayCallout(shape.Callout)
            else
            if (Node.Text = 'Diagram') then
            begin
              try
                if shape.HasDiagram then
                  DisplayDiagram(shape.Diagram);
              except
              end
            end else
            if (Node.Text = 'OLEFormat') then
              DisplayOLEFormat(shape.OLEFormat)
            else
            if (Node.Text = 'PictureFormat') then
              DisplayPictureFormat(shape.PictureFormat)
            else
            if (Node.Text = 'Fill') then
              DisplayFillFormat(shape.Fill)
            else
            if (Node.Text = 'Shadow') then
              DisplayShadow(shape.Shadow)
            else
            if (Node.Text = 'TextEffect') then
              DisplayTextEffect(shape.TextEffect)
            else
            if (Node.Text = 'TextFrame') then
              DisplayTextFrame(shape.TextFrame)
            else
            if (Node.Text = 'GroupItems') then
              DisplayGroupItems(shape.GroupItems);
          end;
    end;
  finally
    ValueList.Items.EndUpdate;
    Screen.Cursor := crs;
  end;
end;

procedure TMainForm.TreeViewExpanding(Sender: TObject; ANode: TTreeNode;
  var AllowExpansion: Boolean);

  function AddChildNode(ANode: TTreeNode; const AText: string;
    AImageIndex: integer; AHasChildren: boolean): TTreeNode;
  begin
    Result := TreeView.Items.AddChild(ANode, AText);
    with Result do begin
      ImageIndex := AImageIndex;
      SelectedIndex := AImageIndex;
      HasChildren := AHasChildren;
    end;
  end;

  procedure AddFirstNode(ANode: TTreeNode);
  begin
    AddChildNode(ANode, 'BuiltInDocumentProperties', 9, false);
    AddChildNode(ANode, 'CustomDocumentProperties', 9, false);
    AddChildNode(ANode, 'DefaultShape', 1, true);
    AddChildNode(ANode, 'Masters', 19, true);
    AddChildNode(ANode, 'PageSetup', 18, false);
    AddChildNode(ANode, 'Slides', 3, FActPresentation.Slides.Count > 0);
    AddChildNode(ANode, 'SlideshowSettings', 17, false);
  end;

  procedure AddHeaderFooter(ANode: TTreeNode; const AName: string);
  begin
    with TreeView.Items.AddChild(ANode, AName) do
    begin
      if AName = 'DateAndTime' then
        ImageIndex := 6
      else if AName='SlideNumber' then
        ImageIndex := 5
      else if AName='Header' then
        ImageIndex := 7
      else if AName='Footer' then
        ImageIndex := 8
      else
        ImageIndex := 4;
      SelectedIndex := ImageIndex;
    end;
  end;

  procedure AddSlides(ASlidesNode: TTreeNode);
  var
    slide: OLEVariant;
    slideNode: TTreeNode;
    i, j: integer;
    s: string;
  begin
    for i := 1 to FActPresentation.Slides.Count do begin
      slide := FActPresentation.Slides.Item(i);
      s := slide.Name;
      j := slide.SlideNumber;
      s := Format('%s (#%d)', [s, j]);
      SlideNode := TreeView.Items.AddChildObject(ASlidesNode, s,
        pointer(PtrInt(Slide.SlideID)));
      with SlideNode do begin
        ImageIndex := 3;
        SelectedIndex := 3;
        HasChildren := (slide.Shapes.Count > 0);
      end;
    end;
  end;

  procedure AddSlideShapes(ANode: TTreeNode);
  var
    slide: OLEVariant;
    shape: OLEVariant;
    shapeNode: TTreeNode;
    i: integer;
  begin
    slide := FindSlide(ANode);
    for i := 1 to slide.Shapes.Count do begin
      Shape := Slide.Shapes.Item(i);
      AddChildNode(ANode, Shape.Name, 1, true);
    end;
  end;

  procedure AddGroupedShapes(ANode: TTreeNode);
  var
    groupshape: OLEVariant;
    shape: OLEVariant;
    n, i: integer;
  begin
    groupshape := FindShape(ANode.Parent);
    if groupshape.&Type = msoGroup then begin
      n := groupshape.GroupItems.Count;
      if n > 0 then begin
        for i := 1 to n do begin
          shape := groupshape.GroupItems.Item(i);
          AddChildNode(ANode, shape.Name, 1, true);
        end;
      end;
    end;
  end;

  procedure AddMaster(AMastersNode: TTreeNode; AMaster: OLEVariant;
    AMasterName: string);
  begin
    AddChildNode(AMastersNode, AMasterName, 19, AMaster.Shapes.Count > 0);
  end;

  procedure AddMasterShapes(ANode: TTreeNode);
  var
    AMaster: OLEVariant;
    shape: OLEVariant;
    shapeNode: TTreeNode;
    i: integer;
    st: integer;
  begin
    if ANode.Parent.Text = 'SlideMaster' then
      AMaster := FActPresentation.SlideMaster
    else if ANode.Parent.Text = 'TitleMaster' then
      AMaster := FActPresentation.Titlemaster
    else if ANode.Parent.Text = 'NotesMaster' then
      AMaster := FActPresentation.NotesMaster
    else if ANode.Parent.Text = 'HandoutMaster' then
      AMaster := FActPresentation.HandoutMaster
    else
      exit;

    for i := 1 to AMaster.Shapes.Count do begin
      shape := AMaster.Shapes.Item(i);
      st := Shape.&Type;
      shapeNode := TreeView.Items.AddChild(ANode, shape.Name);
      with shapeNode do begin
        ImageIndex := 1;
        SelectedIndex := 1;
        HasChildren := (st in [msoPicture, msoLinkedPicture,
          msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject,
          msoTextEffect, msoGroup]);
        if (not hasChildren) and shape.HasTextFrame and shape.TextFrame.HasText then
          HasChildren := true;
      end;
    end;
  end;

  procedure AddShape(ANode: TTreeNode; Shape: OLEVariant);
  // try-except due to  "DefaultShape"
  begin
    try
      if Shape.&Type <> msoGroup then
        AddChildNode(ANode, 'ActionSettings (ppMouseClick)', 23, false);
    except
    end;

    try
      if shape.&Type <> msoGroup then
        AddChildNode(ANode, 'ActionSettings (ppMouseOver)', 23, false);
    except
    end;

    try
      AddChildNode(ANode, 'AnimationSettings', 15, false);
    except
    end;

    try
      AddChildNode(ANode, 'Callout', 24, false);
    except
    end;

    try
      if Shape.HasDiagram then
        AddChildNode(ANode, 'Diagram', 25, false);
    except
    end;

    try
      AddChildNode(ANode, 'Fill', 20, false);
    except
    end;

    try
      if (Shape.&Type = msoGroup) then
        AddChildNode(ANode, 'GroupItems', 1, shape.GroupItems.Count>0);
    except
    end;

    try
      AddChildNode(ANode, 'LinkFormat', 31, false);
    except
    end;
    try
      if (Shape.&Type = msoEmbeddedOLEObject)
        or (Shape.&Type = msoOLEControlObject)
        or (Shape.&Type = msoLinkedOLEObject) then
      begin
        AddChildNode(ANode, 'OLEFormat', 10, false);
      end;
    except
    end;

    try
      if (Shape.&Type = msoPicture) or (Shape.&Type = msoLinkedPicture) then
        AddChildNode(ANode, 'PictureFormat', 11, false);
    except
    end;

    try
      AddChildNode(ANode, 'Shadow', 14, false);
    except
    end;

    try
      if (Shape.&Type = msoTextEffect) then
        AddChildNode(ANode, 'TextEffect', 13, false);
    except
    end;

    try
      if (Shape.HasTextFrame) and (Shape.TextFrame.HasText) then
        AddChildNode(ANode, 'TextFrame', 12, false);
    except
    end;
  end;

  procedure AddSequences(ANode: TTreeNode; ATimeLine: OLEVariant);
  var
    i: integer;
  begin
    AddChildNode(ANode, 'MainSequence', 28, true);
    for i := 1 to ATimeLine.InterActiveSequences.Count do
      AddChildNode(ANode, Format('InteractiveSequence.Item(%d)', [i]), 28, true);
  end;

  procedure AddSequence(ANode: TTreeNode; ATimeLine: OLEVariant);
  var
    seq: OLEVariant;
    nr: string;
    i: integer;
  begin
    if ANode.Text = 'MainSequence' then
      seq := ATimeLine.MainSequence
    else
    begin
      nr := copy(ANode.Text, pos('(', ANode.Text)+1, 255);
      Delete(nr, Length(nr), 1);
      seq := ATimeLine.InteractiveSequences.Item(StrToInt(nr));
    end;
    for i:=1 to seq.Count do
      AddChildNode(ANode, Format('Effect #%d', [i]), 27, true);
  end;

  procedure AddBehaviors(ANode: TTreeNode);
  var
    eff: OLEVariant;
    i: integer;
  begin
    eff := SelectEffect(ANode);
    for i := 1 to eff.Behaviors.Count do
      AddChildNode(ANode, Format('Behaviors.Item(%d)', [i]), 30, false);
  end;

var
  crs: TCursor;
  shape: OLEVariant;
begin
  AllowExpansion := true;
  crs := Screen.Cursor;
  Screen.Cursor := crHourglass;
  TreeView.Items.BeginUpdate;
  try
    FActPresentation := FindPresentation(ANode);
    case ANode.Level of
      0 : AddFirstNode(ANode);
      1 : if (ANode.Text = 'Slides') then
            AddSlides(ANode)
          else
          if (ANode.Text = 'Masters') then
          begin
            AddMaster(ANode, FActPresentation.SlideMaster, 'SlideMaster');
            if FActPresentation.HasTitleMaster then
              AddMaster(ANode, FActPresentation.TitleMaster, 'TitleMaster');
            AddMaster(ANode, FActPresentation.NotesMaster, 'NotesMaster');
            AddMaster(ANode, FActPresentation.HandoutMaster, 'HandoutMaster');
          end else
          if (ANode.Text = 'DefaultShape') then
            AddShape(ANode, FActPresentation.DefaultShape);
      2 : if (ANode.Parent.Text = 'Slides') then
          begin
            AddChildNode(ANode, 'HeadersFooters', 4, true);
            if GetPowerpointVersion(Powerpoint) >= 10 then
              AddChildNode(ANode, 'Sequences', 28, true);
            AddChildNode(ANode, 'Shapes', 1, true);
            AddChildNode(ANode, 'SlideshowTransition', 16, false);
          end else
          if (ANode.Parent.Text = 'Masters') then
          begin
            AddChildNode(ANode, 'HeadersFooters', 4, true);
            if GetPowerpointVersion(Powerpoint) >= 10 then
              AddChildNode(ANode, 'Sequences', 28, true);
            AddChildNode(ANode, 'Shapes', 1, true);
            AddChildNode(ANode, 'SlideshowTransition', 16, false);
          end;
      3 : if (ANode.Text = 'HeadersFooters') then
          begin
            AddHeaderFooter(ANode, 'DateAndTime');
            AddHeaderFooter(ANode, 'Footer');
            AddHeaderFooter(ANode, 'Header');
            AddHeaderFooter(ANode, 'SlideNumber');
          end else
          if (ANode.Text='Shapes') then
          begin
            if ANode.Parent.Parent.Text = 'Slides' then
              AddSlideShapes(ANode)
            else if ANode.Parent.Parent.Text = 'Masters' then
              AddMasterShapes(ANode);
          end else
          if (ANode.Text = 'Sequences') then
          begin
            if ANode.Parent.Parent.Text = 'Slides' then
              AddSequences(ANode, FindSlide(ANode).TimeLine)
            else if (ANode.Parent.Parent.Text = 'SlideMaster') then
              AddSequences(ANode, FActPresentation.SlideMaster.TimeLine);
          end;
      4 : if (ANode.Parent.Text = 'Shapes') then
          begin
            shape := FindShape(ANode);
            AddShape(ANode, shape);
          end else
          if (ANode.Parent.Text = 'Sequences') then begin
            if (ANode.Parent.Parent.Parent.Text = 'Slides') then
              AddSequence(ANode, FindSlide(ANode).TimeLine)
            else if (ANode.Parent.Parent.Parent.Text = 'SlideMaster') then
              AddSequence(ANode, FActPresentation.SlideMaster.TimeLine);
          end;
      5 : if (ANode.Text = 'GroupItems') then
            AddGroupedShapes(ANode)
          else
          if (ANode.Parent.Parent.Text = 'Sequences') then
          begin
            AddBehaviors(ANode);
            AddChildNode(ANode, 'EffectInformation', 29, false);
          end;
      6 : if (ANode.Parent.Text = 'GroupItems') then
          begin
            shape := FindShape(ANode);
            AddShape(ANode, shape);
          end;
    end;
  finally
    Screen.Cursor := crs;
    TreeView.Items.EndUpdate;
  end;
end;

procedure TMainForm.ValueListDblClick(Sender: TObject);
var
  F: TPropertyForm;
  item: TListItem;
begin
  F := TPropertyForm.Create(nil);
  try
    item := ValueList.Selected;
    F.PropertyName := item.Caption;
    F.PropertyValue := item.SubItems[0];
    F.DisplayOption := item.SubItems[1];
    F.ShowModal;
  finally
    F.Free;
  end;
end;

function TMainForm.IsDirty(ANode: TTreeNode): boolean;
begin
  ANode := FindPresentationNode(ANode);
  Result := PtrInt(ANode.Data) > 0;
end;

function TMainForm.NodeIsShape(ANode: TTreeNode): boolean;
begin
  Result := Assigned(ANode) and Assigned(ANode.Parent) and (ANode.Parent.Text = 'Shapes');
end;

function TMainForm.NodeIsSlide(ANode: TTreeNode): boolean;
begin
  Result := Assigned(ANode) and Assigned(ANode.Parent) and (ANode.Parent.Text = 'Slides');
end;

procedure TMainForm.PopulateTreeView(APresentationIndex: Integer);
var
  PresNode: TTreeNode;
  crs: TCursor;
begin
  crs := Screen.Cursor;
  Screen.Cursor := crHourglass;
  TreeView.Items.BeginUpdate;
  try
    PresNode := TreeView.Items.AddObject(
      nil,
      ExtractFileName(FFilename),
      TObject(PtrInt(APresentationIndex))
    );
    with PresNode do begin
      ImageIndex := 0;
      SelectedIndex := 0;
      HasChildren := true;
      Expanded := true;
      Selected := true;
    end;
  finally
    TreeView.Items.EndUpdate;
    Screen.Cursor := crs;
  end;
end;

procedure TMainForm.ReadIni;
var
  ini: TCustomIniFile;
begin
  ini := CreateIni;
  try
    cbShowerrorItems.Checked := ini.ReadBool('MainForm', 'ShowErrorItems', cbShowErrorItems.Checked);
  finally
    ini.Free;
  end;
end;

function TMainForm.SelectEffect(ANode: TTreeNode): OLEVariant;
var
  seq: OLEVariant;
  nr: string;
begin
  seq := SelectSequence(ANode.Parent);
  nr := copy(ANode.Text, pos('#', ANode.Text)+1, 255);
  Result := seq.Item(StrToInt(nr));
end;

function TMainForm.SelectHeaderFooter(ANode: TTreeNode): OLEVariant;
var
  ownr: OLEVariant;
begin
  Result := Unassigned;
  if ANode.Parent.Parent.Parent.Text = 'Slides' then
    ownr := SelectSlide(ANode.Parent.Parent)
  else if ANode.Parent.Parent.Parent.Text = 'Masters' then
    ownr := SelectMaster(ANode.Parent.Parent)
  else
    exit;

  if VarIsEmpty(ownr)
    then exit;

  if ANode.Text = 'Footer'
    then Result := ownr.HeadersFooters.Footer
  else if ANode.Text = 'Header'
    then Result := ownr.HeadersFooters.Header
  else if ANode.Text = 'DateAndTime'
    then Result := ownr.HeadersFooters.DateAndTime
  else if ANode.Text = 'SlideNumber'
    then Result := ownr.HeadersFooters.SlideNumber;
end;

function TMainForm.SelectMaster(ANode: TTreeNode): OLEVariant;
var
  win: OLEVariant;
begin
  Result := UnAssigned;
  SelectPresentation(ANode);

  if ANode.Text = 'SlideMaster' then
    Result := FActPresentation.SlideMaster
  else if ANode.Text = 'TitleMaster' then
    Result := FActPresentation.TitleMaster
  else if ANode.Text = 'NotesMaster' then
    Result := FActPresentation.NotesMaster
  else if ANode.Text = 'HandoutMaster' then
    Result := FActPresentation.HandoutMaster;

  try
    win := FPowerpoint.ActiveWindow;
    if ANode.Text = 'SlideMaster' then
      win.ViewType := ppViewSlideMaster
    else if ANode.Text='TitleMaster' then
      win.ViewType := ppViewTitleMaster
    else if ANode.Text='NotesMaster' then
      win.ViewType := ppViewNotesMaster
    else if ANode.Text='HandoutMaster' then
      win.ViewType := ppViewHandoutMaster;
  except
    // in case of PPT not being open...
  end;
end;

function TMainForm.SelectPresentation(ANode: TTreeNode): OLEVariant;
begin
  Result := FindPresentation(ANode);
  if CompareText(Result.FullName, FActPresentation.Fullname) <> 0 then begin
    FActPresentation := Result;
    FActPresentation.Windows.Item(1).Activate;
    FActSlideID := -1;
  end;
end;

function TMainForm.SelectSequence(ANode: TTreeNode): OLEVariant;
var
  slide: OLEVariant;
  nr: string;
  TimeLine: OLEVariant;
begin
  if ANode.Parent.Parent.Parent.Text = 'Slides' then
  begin
    slide := SelectSlide(ANode.Parent.Parent);
    TimeLine := slide.TimeLine;
  end else
  begin
    SelectPresentation(ANode);
    TimeLine := FActPresentation.SlideMaster.TimeLine;
  end;

  if ANode.Text = 'MainSequence' then
    Result := TimeLine.MainSequence
  else
  begin
    nr := copy(ANode.Text, pos('(', ANode.Text)+1, 255);
    Delete(nr, Length(nr), 1);
    Result := TimeLine.InteractiveSequences.Item(StrToInt(nr));
  end;
end;

function TMainForm.SelectShape(ANode: TTreeNode): OLEVariant;
begin
  Result := UnAssigned;

  if (ANode = nil)
    or (ANode.Parent = nil)
    or (ANode.Parent.Parent = nil)
    or (ANode.Parent.Parent.Parent = nil)
  then
    exit;

  if (ANode.Parent.Parent.Parent.Text = 'Slides') then
    SelectSlide(ANode.Parent.Parent)
  else if (ANode.Parent.Parent.Parent.Parent <> nil)
      and (ANode.Parent.Parent.Parent.Parent.Parent <> nil)
      and (ANode.Parent.Parent.Parent.Parent.Parent.Text = 'Slides')
  then
    SelectSlide(ANode.Parent.Parent.Parent.Parent)
  else
  if   (ANode.Parent.Parent.Text = 'SlideMaster')
    or (ANode.Parent.Parent.Text = 'TitleMaster')
    or (ANode.Parent.Parent.Text = 'HandoutMaster')
    or (ANode.Parent.Parent.Text = 'NotesMaster')
  then
    SelectMaster(ANode.Parent.Parent)
  else
    exit;

  Result := FindShape(ANode);
  try
    if not VarIsEmpty(Result) then
      Result.Select;
  except
    // in case that PPT is not open...
  end;
end;

function TMainForm.SelectSlide(ANode: TTreeNode): OLEVariant;
var
  win: OLEVariant;
begin
  Result := Unassigned;
  if (ANode.Parent.Text = 'Slides')
    or ( (ANode.Parent.Parent <> nil)
         and (ANode.Parent.Parent <> nil)
         and (ANode.Parent.Parent.Parent <> nil)
         and (ANode.Parent.Parent.Parent.Text = 'Slides')
       )
  then begin
    SelectPresentation(ANode);
    result := FindSlide(ANode.Text);
    if not VarIsEmpty(Result)
      and ((FActSlideID = -1) or (Result.SlideID <> FActSlideID))
    then begin
      FActSlideID := Result.SlideID;
      try
        win := PowerPoint.ActiveWindow;
        win.ViewType := ppViewSlideSorter;
        Result.Select;
        win.ViewType := ppViewSlide;
      except
        // in case Powerpoint has not been opened.
      end;
    end;
  end;
end;

procedure TMainForm.SetDirty(ANode: TTreeNode; AValue: boolean);
var
  already_dirty: boolean;
  idx: PtrInt;
begin
  ANode := FindPresentationNode(ANode);
  idx := abs(PtrInt(ANode.Data));
  if AValue then
    ANode.Data := pointer(idx)
  else
    ANode.Data := pointer(-idx);

  already_dirty := copy(ANode.Text, Length(ANode.Text)-3, 3) = '[*]';

  if not AValue and already_dirty then
    ANode.Text := trim(copy(ANode.Text, 1, Length(ANode.Text)-3))
  else
  if AValue and (not already_dirty) then
    ANode.Text := ANode.Text + ' [*]';
end;

procedure TMainForm.WriteIni;
var
  ini: TCustomIniFile;
begin
  ini := CreateIni;
  try
    ini.WriteBool('MainForm', 'ShowErrorItems', cbShowErrorItems.Checked);
  finally
    ini.Free;
  end;
end;

end.

