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

uses vcl.stdCtrls, System.SysUtils, uOpenOffice_calc, uOpenOffice_writer, uOpenOfficeCollors, math,
  System.Variants;

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

  TTypeChart = (ctDefault, ctVertical, ctPie, ctLine);

  THelperHoriJustify = record helper for THoriJustify
  public
    function toInteger: Integer;
  end;

  THelperVertJustify = record helper for TVertJustify
  public
    function toInteger: Integer;
  end;

  TSettingsChart = record
    Height,
    Width,
    Position_X,
    Position_Y,
    StartRow,
    PositionSheet,
    EndRow: integer;
    StartColumn,
    EndColumn,
    ChartName: string;
    typeChart: TTypeChart;
  end;

  THelperOpenOffice_writer = class helper for TOpenOffice_writer
    function setUnderline(aUnderline: boolean): TOpenOffice_writer;
    function setBold(aBold: boolean): TOpenOffice_writer;
    function setFontHeight(aFontHeight: integer): TOpenOffice_writer;
    function setColorText(aColor: TOpenColor) : TOpenOffice_writer;
    function setFontName(aFont : string): TOpenOffice_writer;
  end;

  THelperOpenOffice_calc = class helper for TOpenOffice_calc
    procedure addChart(aSettingsChart: TSettingsChart);
    function setBorder(borderPosition: TBoderSheet; opColor: TOpenColor; RemoveBorder: boolean = false) : TOpenOffice_calc;
    function changeFont(aNameFont: string; aHeight: Integer): TOpenOffice_calc;
    function changeJustify(aTypeHori: THoriJustify; aTypeVert: TVertJustify) : TOpenOffice_calc;
    function setColor(aFontColor, aBackgroud: TOpenColor): TOpenOffice_calc;
    function setCellWidth(const aWidth: integer): TOpenOffice_calc;
    function setBold(aBold: boolean): TOpenOffice_calc;
    function SetUnderline(aUnderline: boolean): TOpenOffice_calc;
    function CountRow: Integer;
    function CountCell: Integer;
  end;

implementation

procedure THelperOpenOffice_calc.addChart(aSettingsChart: TSettingsChart);
var
  Chart, Rect, sheet : OleVariant;
  RangeAddress: Variant;
  countChart: Integer;
begin
  countChart := 1;

  if aSettingsChart.ChartName.trim.IsEmpty then
    aSettingsChart.ChartName := 'MyChart_' + (aSettingsChart.StartColumn + aSettingsChart.StartRow.ToString) + '_' +
      (aSettingsChart.EndColumn + aSettingsChart.EndRow.ToString);

  sheet := objDocument.Sheets.getByIndex(aSettingsChart.PositionSheet);
  // getByName(aCollName);
  Charts := sheet.Charts;

  while Charts.hasByName(aSettingsChart.ChartName) do
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

function THelperOpenOffice_calc.changeFont(aNameFont: string; aHeight: Integer)
  : TOpenOffice_calc;
begin
  // Cell := Table.getCellRangeByName(aCollName+aCellNumber.ToString);
  Cell.CharFontName := aNameFont;
  Cell.CharHeight := inttostr(aHeight);
  result := self;
end;

function THelperOpenOffice_calc.changeJustify(aTypeHori: THoriJustify;
  aTypeVert: TVertJustify): TOpenOffice_calc;
begin
  Cell.HoriJustify := aTypeHori.toInteger;
  Cell.VertJustify := aTypeVert.toInteger;
  result := self;
end;

function THelperOpenOffice_calc.CountRow: Integer;
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

function THelperOpenOffice_writer.setBold(aBold: boolean): TOpenOffice_writer;
var CtrlBold: boolean;
begin
  CtrlBold := false;
  propsText[0].Name := 'bold';
  propsText[0].Value := aBold;
  //Funcionando, porém rever
  if BoldActive and (not aBold) then
  begin
    BoldActive   := false;
    CtrlBold     := true;
    aBold        := true;
  end;

  if ( not BoldActive) and (aBold) then
  begin
    objDispatcher.executeDispatch(objWriter, '.uno:Bold', '', 0,  VarArrayOf(propsText));

    if not CtrlBold then
      BoldActive := true;
  end;

  Result := self;
end;

function THelperOpenOffice_writer.setFontName(aFont: string): TOpenOffice_writer;
begin
  if not ValueText.trim.IsEmpty then
  begin
    Cursor.CharFontName := aFont;
    setValue(ValueText);
  end;
  Result := self;
end;

function THelperOpenOffice_writer.setColorText(aColor: TOpenColor): TOpenOffice_writer;
begin
  if not ValueText.Trim.IsEmpty then
  begin
     Cursor.SetPropertyValue('CharColor', aColor);
     setValue(ValueText);
  end;

  Result := self;
end;

function THelperOpenOffice_writer.setFontHeight(aFontHeight: integer) : TOpenOffice_writer;
begin
  propsText[1].Name := 'FontHeight.Height';
  propsText[1].Value := aFontHeight;
  objDispatcher.executeDispatch(objWriter, '.uno:FontHeight', '', 0, VarArrayOf(propsText));

  Result := self;
end;

function THelperOpenOffice_writer.setUnderline(aUnderline: boolean): TOpenOffice_writer;
begin
  if not ValueText.Trim.IsEmpty then
  begin
    Cursor.CharUnderline := ifthen(aUnderline,1,0);
    setValue(ValueText);
  end;

  Result := self;
end;

function THelperOpenOffice_calc.CountCell: Integer;
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

function THelperOpenOffice_calc.seTBorder(borderPosition: TBoderSheet; opColor: TOpenColor; RemoveBorder: boolean): TOpenOffice_calc;
var
  border: Variant;
  settings: Variant;
begin
  border := ServicesManager.createInstance('com.sun.star.reflection.CoreReflection');
  border.forName('com.sun.star.table.BorderLine2').createObject(settings);

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

function THelperOpenOffice_calc.setCellWidth(const aWidth: integer): TOpenOffice_calc;
begin
   Cell.getColumns.getByIndex(0).Width := aWidth;
end;

function THelperOpenOffice_calc.setColor(aFontColor, aBackgroud: TOpenColor)
  : TOpenOffice_calc;
begin
  Cell.CharColor := aFontColor;
  Cell.CellBackColor := aBackgroud;
  result := self;
end;

function THelperOpenOffice_calc.setBold(aBold: boolean): TOpenOffice_calc;
begin
  Cell.CharWeight := ifthen(aBold, 150, 0);
  result := self;
end;

function THelperOpenOffice_calc.SetUnderline(aUnderline: boolean): TOpenOffice_calc;
begin
  Cell.CharUnderline := ifthen(aUnderline, 1, 0);
  result := self;
end;

{ THelperOpenOffice_calc }

function THelperHoriJustify.toInteger: Integer;
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

function THelperVertJustify.toInteger: Integer;
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
