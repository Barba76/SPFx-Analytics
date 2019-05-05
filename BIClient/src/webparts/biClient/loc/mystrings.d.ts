declare interface IBiClientWebPartStrings {
  DataSourceLabel: string;
  TitleLabel: string;
  WidgetLabel: string;
  WidgetTypeLabel: string;
  ChartOption: string;
  PieChartOption: string;
  KpiCardOption: string;
  ChartType: string;
  BarsSingleOption: string;
  BarsDoubleOption: string;
  LinesSingleOption: string;
  LinesSoubleOption: string;
  LabelsLabel: string;
  MainValueLabel: string;
  SecondaryValueLabel: string;
  MaxHeightLabel: string;
  MaxHeightDescription: string;
  InvalidValueMessage: string;
  MinValueMessage: string;
}

declare module 'BiClientWebPartStrings' {
  const strings: IBiClientWebPartStrings;
  export = strings;
}
