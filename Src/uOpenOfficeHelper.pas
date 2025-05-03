{ ******************************************************* }

{ Delphi openOffice Library }

{ File     : uOpenOfficeHelper.pas }
{ Developer: Daniel Fernandes Rodrigures }
{ Email    : danielfernandesroddrigues@gmail.com }
{ this unit is a part of the Open Source. }
{ licensed under a MIT - see LICENSE.md}

{ ******************************************************* }

unit uOpenOfficeHelper;

interface

uses
  Vcl.stdCtrls,
  System.SysUtils,
  math,
  System.Variants,
  uOpenOffice_calc,
  uOpenOfficeCollors,
  uOpenOffice.Chart,
  uOpenOffice.DrawImage;

type

  TBorder = (bAll, bLeft, bRight, bBottom, bTop);

  TBoderSheet = set of TBorder;

  { STANDARD : é o alinhamento padrão tanto para números como para textos, sendo a esqueda para as strings e a direita para os números;
    LEFT : o conteúdo é alinhado no lado esquerdo da célula;
    CENTER : o conteúdo é alinhado no centro da célula;
    RIGHT : o conteúdo é alinhado no lado direito da célula;
    BLOCK : o conteúdo é alinhando em relação ao comprimento da célula;
    REPEAT : o conteúdo é repetido dentro da célula para preenchê-la. }
  THoriJustify = (fthSTANDARD, fthLEFT, fthCENTER, fthRIGHT, fthBLOCK,
    fthREPEAT);
  { STANDARD : é o valor usado como padrão;
    TOP : o conteúdo da célula é alinhado pelo topo;
    CENTER : o conteúdo da célula é alinhado pelo centro;
    BOTTOM : o conteúdo da célula é alinhado pela base. }
  TVertJustify = (ftvSTANDARD, ftvTOP, ftvCENTER, ftvBOTTOM);

  TTypeChart = TypeChart;
  TSettingsChart = TChart;

  THelperHoriJustify = record helper for THoriJustify
  public
    function ToInteger: Integer;
  end;

  THelperVertJustify = record helper for TVertJustify
  public
    function ToInteger: Integer;
  end;

  THelperOpenOffice_calc = class helper for TOpenOffice_Calc
    procedure AddChart(aSettingsChart: TSettingsChart);
    function SetBorder(borderPosition: TBoderSheet; opColor: TOpenColor; RemoveBorder: boolean = false) : TOpenOffice_Calc;
    function ChangeFont(aNameFont: string; aHeight: Integer): TOpenOffice_Calc;
    function ChangeJustify(aTypeHori: THoriJustify; aTypeVert: TVertJustify) : TOpenOffice_Calc;
    function SetColor(aFontColor, aBackgroud: TOpenColor): TOpenOffice_Calc;
    function SetCellWidth(const aWidth: integer): TOpenOffice_Calc;
    function SetBold(aBold: boolean): TOpenOffice_Calc;
    function SetUnderline(aUnderline: boolean): TOpenOffice_Calc;
    function DrawImage(aImage : TOpenOfficeDrawImage) :TOpenOffice_Calc;
    function CountRow: Integer;
    function CountCell: Integer;
    function SheetToBase64(aPathFile:string):string;
  end;

implementation

uses
  System.Win.ComObj, System.Classes, Soap.EncdDecd;

procedure THelperOpenOffice_calc.AddChart(aSettingsChart: TSettingsChart);
var
  Chart, Rect, Sheet : OleVariant;
  RangeAddress: Variant;
  CountChart: Integer;
begin
  countChart := 1;

  if aSettingsChart.ChartName.Trim.IsEmpty then
    aSettingsChart.ChartName := 'MyChart_' + (aSettingsChart.StartColumn + aSettingsChart.StartRow.ToString) + '_' +
      (aSettingsChart.EndColumn + aSettingsChart.EndRow.ToString);

  sheet := objDocument.Sheets.getByIndex(aSettingsChart.PositionSheet);
  // getByName(aCollName);
  Charts := sheet.Charts;

  while Charts.HasByName(aSettingsChart.ChartName) do
  begin
    aSettingsChart.ChartName := copy(aSettingsChart.ChartName,0, ifthen( (pos('_',aSettingsChart.ChartName) > 0),
                                               pos('_',aSettingsChart.ChartName), aSettingsChart.ChartName.Length)
                      ) + '_' + countChart.ToString;
    inc(countChart);
    aSettingsChart.Position_Y := (aSettingsChart.Position_Y + aSettingsChart.Height) + 1000;
  end;

  Rect := objServiceManager.Bridge_GetStruct('com.sun.star.awt.Rectangle');
  RangeAddress := sheet.Bridge_GetStruct('com.sun.star.table.CellRangeAddress');

  Rect.Width := aSettingsChart.Width;
  Rect.Height := aSettingsChart.Height;
  Rect.X := aSettingsChart.Position_X;
  Rect.Y := aSettingsChart.Position_Y;

  RangeAddress.sheet := aSettingsChart.PositionSheet;
  RangeAddress.StartColumn := Fields.getIndex(aSettingsChart.StartColumn);
  RangeAddress.StartRow := aSettingsChart.StartRow;
  RangeAddress.EndColumn := Fields.getIndex(aSettingsChart.EndColumn);
  RangeAddress.EndRow := aSettingsChart.EndRow;

  Charts.addNewByName(aSettingsChart.ChartName, Rect, VarArrayOf(RangeAddress), true, true);

  if aSettingsChart.typeChart <> ctDefault then
  begin
    Chart := Charts.getByName(aSettingsChart.ChartName).embeddedObject;
    Chart.Title.String := aSettingsChart.ChartName;

    case aSettingsChart.typeChart of
      ctVertical:
        Chart.Diagram.Vertical := true;
      ctPie:
        begin
          Chart.Diagram := Chart.createInstance
            ('com.sun.star.chart.PieDiagram');
          Chart.HasMainTitle := true;
        end;
      ctLine:
        begin
          Chart.Diagram := Chart.createInstance
            ('com.sun.star.chart.LineDiagram');
        end;
    end;
  end;
end;

function THelperOpenOffice_Calc.ChangeFont(aNameFont: string; aHeight: Integer): TOpenOffice_calc;
begin
  if not aNameFont.Trim.IsEmpty then
    Cell.CharFontName := aNameFont;

  Cell.CharHeight := inttostr(aHeight);
  result := self;
end;

function THelperOpenOffice_Calc.ChangeJustify(aTypeHori: THoriJustify;
  aTypeVert: TVertJustify): TOpenOffice_calc;
begin
  Cell.HoriJustify := aTypeHori.toInteger;
  Cell.VertJustify := aTypeVert.toInteger;
  result := self;
end;

function THelperOpenOffice_Calc.CountRow: Integer;
var
  FRow, FCountRow: Integer;
  FCountBlank: Integer;
  FBreak, allBlank: boolean;
  I: Integer;
begin
  FBreak := false;
  FRow := 1;
  FCountRow := 0;
  FCountBlank := 0;

  while not FBreak do
  begin
    for I := 0 to 21 do
    begin
      if GetValue(FRow, Fields.getField(I)).Value.trim.IsEmpty then
      begin
        allBlank := true;
      end
      else
      begin
        if FCountBlank > 0 then // An empty column behind a valued column
          FCountRow := FCountRow + FCountBlank;

        allBlank := false;
        FCountBlank := 0;

        inc(FCountRow);
        break;
      end;
    end;
    inc(FRow);

    if FCountBlank = 50 then
      FBreak := true;

    if allBlank then
      inc(FCountBlank);

  end;
  result := FCountRow;
end;

function THelperOpenOffice_calc.DrawImage(aImage: TOpenOfficeDrawImage): TOpenOffice_Calc;
var
  DrawPage: OleVariant;
begin
  if not Assigned(aImage) then
    raise Exception.Create('Error, the object is not created');

  DrawPage := oSCalc.getDrawPage;
  Image := ServicesManager.CreateInstance('com.sun.star.drawing.GraphicObjectShape');
  try
    Image.GraphicURL := 'file:///' + StringReplace(aImage.URLImage, '\', '/', [rfReplaceAll]);
    Image.setPosition(Int(aImage.PositionX), Int(aImage.PositionY));
    Image.setSize(Int(aImage.Width), Int(aImage.Height));

    DrawPage.add(Image);
  finally
    Image := Unassigned;
    Result := Self;
  end;
end;

function THelperOpenOffice_Calc.CountCell: Integer;
var
  FCell, FCountCell, FCountBlank: Integer;
  I: Integer;
  allBlank: boolean;
begin
  FCell := 1;
  FCountCell := 0;
  FCountBlank := 0;

  for I := 0 to 21 do
  begin
    for FCell := 1 to 10 do
    begin
      if not GetValue(FCell, Fields.getField(I)).Value.trim.IsEmpty then
      begin

        if FCountBlank > 0 then
          FCountCell := FCountCell + FCountBlank;

        allBlank := false;

        inc(FCountCell);
        break;
      end
      else
        allBlank := true;
    end;

    if FCountBlank = 10 then
    begin
      FCountBlank := 0;
      break;
    end;

    if allBlank then
      inc(FCountBlank);
  end;

  result := FCountCell;
end;

function THelperOpenOffice_Calc.SetBorder(borderPosition: TBoderSheet; opColor: TOpenColor; RemoveBorder: boolean): TOpenOffice_calc;
var
  settings: Variant;
begin
 CoreReflection.forName('com.sun.star.table.BorderLine2').createObject(settings);

 if not RemoveBorder then
  begin
    settings.Color := opColor;
    settings.InnerLineWidth := 20;
    settings.LineDistance := 60;
    settings.LineWidth := 2;
    settings.OuterLineWidth := 20;
  end else
  begin
    settings.Color := 0;
    settings.InnerLineWidth := 0;
    settings.LineDistance := 0;
    settings.LineWidth := 0;
    settings.OuterLineWidth := 0;
  end;

  if bAll in borderPosition then
  begin
    Cell.TopBorder := settings;
    Cell.LeftBorder := settings;
    Cell.RightBorder := settings;
    Cell.BottomBorder := settings;
  end;

  if bTop in borderPosition then
    Cell.TopBorder := settings;

  if bLeft in borderPosition then
    Cell.LeftBorder := settings;

  if bRight in borderPosition then
    Cell.RightBorder := settings;

  if bBottom in borderPosition then
    Cell.BottomBorder := settings;

  result := self;
end;

function THelperOpenOffice_calc.SetCellWidth(const aWidth: integer): TOpenOffice_calc;
begin
   Cell.getColumns.getByIndex(0).Width := aWidth;
end;

function THelperOpenOffice_calc.SetColor(aFontColor, aBackgroud: TOpenColor): TOpenOffice_calc;
begin
  Cell.CharColor := aFontColor;
  Cell.CellBackColor := aBackgroud;
  result := self;
end;

function THelperOpenOffice_calc.SetBold(aBold: boolean): TOpenOffice_calc;
begin
  Cell.CharWeight := ifthen(aBold, 150, 0);
  result := self;
end;

function THelperOpenOffice_calc.SetUnderline(aUnderline: boolean): TOpenOffice_calc;
begin
  Cell.CharUnderline := ifthen(aUnderline, 1, 0);
  result := self;
end;

function THelperOpenOffice_calc.SheetToBase64(aPathFile:string): string;
var
  stream: TMemoryStream;
begin
  stream := TMemoryStream.Create;
  try
    stream.LoadFromFile(aPathFile);
    Result := EncodeBase64(stream.Memory, stream.Size);
  finally
    stream.Free;
  end;
end;

{ THelperOpenOffice_calc }

function THelperHoriJustify.ToInteger: Integer;
begin
  case self of
    fthSTANDARD:
      result := 0;
    fthLEFT:
      result := 1;
    fthCENTER:
      result := 2;
    fthRIGHT:
      result := 3;
    fthBLOCK:
      result := 4;
    fthREPEAT:
      result := 5;
    else
      raise Exception.Create('Unknown Justification Value');
  end;
end;

{ THelperVertJustify }

function THelperVertJustify.ToInteger: Integer;
begin
  case self of
    ftvSTANDARD:
      result := 0;
    ftvTOP:
      result := 1;
    ftvCENTER:
      result := 2;
    ftvBOTTOM:
      result := 3;
    else
      raise Exception.Create('Unknown Justification Value');
  end;
end;

end.
