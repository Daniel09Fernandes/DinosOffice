unit uOpenOffice.Chart;

interface

type
 TypeChart = (ctDefault, ctVertical, ctPie, ctLine);

 TChart = record
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
    typeChart: TypeChart;
  end;

implementation

end.
