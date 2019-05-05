import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  DynamicDataSharedDepth,
  IPropertyPaneDropdownOption,
  IPropertyPaneGroup,
  IPropertyPaneField
} from '@microsoft/sp-property-pane';
import { escape, times } from '@microsoft/sp-lodash-subset';

import styles from './BiClientWebPart.module.scss';
import * as strings from 'BiClientWebPartStrings';
import { DynamicProperty } from '@microsoft/sp-component-base';
import { DataSetColumnType, DataSet } from './Model';

export enum WidgetType  {Chart = 1, Pie=2, KPI = 3}
export enum ChartType { SingleBars=1, DoubleBars=2, SingleLines=3, DoubleLines=4}

export interface IBiClientWebPartProps {
  title: string;
  label: string;
  value: string;
  secondaryValue: string;
  bi_dataset: DynamicProperty<Object>;
  widget : WidgetType;
  chartType:ChartType;
  maxHeight : number;
}

var Chart : any = require("chart.js");

export default class BiClientWebPart extends BaseClientSideWebPart<IBiClientWebPartProps>  {

  private static DATA_KEY : string = "bi_dataset";
  private static MIN_HEIGHT : number = 200;

  private widgetData : any;
  private dataSet : any;
  private lastWidgetTitle : string;

  protected onInit(): Promise<void> {   

    if (this.properties.bi_dataset){
      try{
        this.dataSet = this.properties.bi_dataset.tryGetValue();
      }catch(error){
        console.error(error);
      }
    }

    return Promise.resolve();
  }


  public render(): void {


    if (this.properties.bi_dataset){
      this.dataSet = this.properties.bi_dataset.tryGetValue();
    }

    if (!this.renderedOnce || this.properties.title == this.lastWidgetTitle){
      this.domElement.innerHTML = `<div class="${styles.biClient}"><div class="${styles.widget_title}">${this.properties.title}</div><div class="widget-zone"></div></div>`;
      
      if (this.dataSet){

        switch(this.properties.widget){
          case WidgetType.Pie:
            this.widgetData = this.getValuesForPie();
            this.drawPie();
          break;

          case WidgetType.Chart:
          this.widgetData = this.getValuesForChart();
            this.drawChart();
          break;

          case WidgetType.KPI:
          this.widgetData = this.getValuesForKPICard();
            this.drawKPICard();
          break;

          default:
            this.widgetData = this.getValuesForPie();
            this.drawPie();
          break;
        }
      }
    } else if (this.properties.title != this.lastWidgetTitle){

      let title = this.domElement.getElementsByClassName(styles.widget_title)[0];
      if (title) title.innerHTML = this.properties.title;

    }

    this.lastWidgetTitle = this.properties.title;
    
  }

  private drawPie(){
    //try{

    let data = this.widgetData;    
    let lineColor = this.getThemeColor("white");
    let size = this.domElement.clientWidth + 10;
    let height = size <= this.properties.maxHeight ? size : this.properties.maxHeight;
    
    let widgetZone = this.domElement.getElementsByClassName("widget-zone")[0];
    widgetZone.innerHTML = '<canvas id="myChart" width="' + size + '" height="' + height + '"></canvas>';

    let chart = new Chart(widgetZone.getElementsByTagName('canvas')[0], {
      type: 'pie',
      data:{
        datasets: [{
            data: data.data,
            backgroundColor: [
              'rgba(54, 162, 235, 0.4)',
              'rgba(100, 255, 200, 0.4)',
              'rgba(255, 206, 86, 0.4)',
              'rgba(255, 99, 132, 0.4)',
              'rgba(255, 159, 64, 0.4)',
              'rgba(153, 102, 255, 0.4)',
              'rgba(54, 162, 235, 0.4)',
              'rgba(100, 255, 200, 0.4)',
              'rgba(255, 206, 86, 0.4)',
              'rgba(255, 99, 132, 0.4)',
              'rgba(255, 159, 64, 0.4)',
              'rgba(153, 102, 255, 0.4)'
            ],
            borderColor: lineColor
        }],
    
        // These labels appear in the legend and in the tooltips when hovering different arcs
        labels: data.labels
      }        
    });
  }

  private drawKPICard(){

    let goal : number = this.widgetData.goal;
    let value : number = this.widgetData.value;
    let performance : number =  this.widgetData.performance;

    let performanceClass = performance < 1 ? styles.down_value : styles.up_value;
    let performanceIcon = performance < 1 ? "ms-Icon ms-Icon--CaretSolidDown" : "ms-Icon ms-Icon--CaretSolidUp";

    let cardHTML = `
      <div class="${styles.card_widget}">
        <div class="${styles.bi_card} ${styles.kpi_card} ${performanceClass}">
            <div class="${styles.kpi_performance_wrapper}">
              <i class="${styles.kpi_performance_icon} ${performanceIcon}"></i>
            </div>
            <div class="${styles.kpi_info_wrapper}">
              <div class="${styles.kpi_value}"><label>${value.toLocaleString("en-us")}</label><label class="${styles.kpi_label}">Value</label></div>
              <div class="${styles.secondary_information}">
                <div class="${styles.kpi_goal}"><label>${goal.toLocaleString("en-us")}</label><label class="${styles.kpi_label}">Goal</label></div>
                <div class="${styles.kpi_performance}"><label>${(Math.round(performance*10000)/100).toLocaleString("en-us")}%</label><label class="${styles.kpi_label}">Performance</label></div>
              </div>
            </div>
        </div>
      </div>
    `;

    this.domElement.getElementsByClassName("widget-zone")[0].innerHTML = cardHTML;

  }

  private drawChart(){

    let data = this.widgetData;
    let size = this.domElement.clientWidth + 10;
    let height = size <= this.properties.maxHeight ? size : this.properties.maxHeight;
    this.domElement.getElementsByClassName("widget-zone")[0].innerHTML = '<canvas id="myChart" width="' + size + '" height="' + height + '"></canvas>';

    Chart.defaults.global.defaultFontColor = this.getThemeColor("neutralDark") || "#666";
  
    let type = this.properties.chartType == ChartType.DoubleBars || this.properties.chartType==ChartType.SingleBars ? "bar" : "line";

    let dataSets = [];
    dataSets.push({
      label: data.data.legend_a,
      data: data.data.dataset_a,
      backgroundColor: 'rgba(54, 162, 235, 0.4)',
      borderColor: 'rgba(54, 162, 235, 1)',
      borderWidth: 1,
      lineTension: 0 // default is 0.4 for smooth curves
    });


    let lineColor = this.getThemeColor("neutralLight") || "#aaa";
    let lineBorderColor = this.getThemeColor("neutralTertiary") || "#666";
    if (data.data.dataset_b.length){
      dataSets.push({
        label: data.data.legend_b,
        data: data.data.dataset_b,
        backgroundColor: lineColor,
        borderColor: lineBorderColor,
        borderWidth: 1,
        type: type,
        lineTension: 0
      });
    }


   let chart = new Chart(this.domElement.getElementsByTagName('canvas')[0], {
      type: type,
      data: {
          labels: data.data.labels,
          datasets: dataSets,
      },
      options: {
        scales: {
            yAxes: [{
                ticks: {
                    beginAtZero: true
                }
            }]
        },
        legend: {
          display: true
        }
      }
  });



  }

  protected onAfterResize(){
    if (this.properties.widget == WidgetType.Pie){
      this.drawPie();
    } else if (this.properties.widget == WidgetType.Chart) {
      this.drawChart();
    } else if (this.properties.widget == WidgetType.KPI){
      this.drawKPICard();
    }
    
  }

  private getValuesForChart(){
    let doubleChart : boolean = this.properties.chartType == ChartType.DoubleBars || this.properties.chartType == ChartType.DoubleLines;
    let tempResults_a = {};
    let tempResults_b = {};
    let finalResults = {
      data:{
        labels:[],
        legend_a: this.dataSet.columns[this.properties.value].name,
        legend_b: doubleChart ? this.dataSet.columns[this.properties.secondaryValue].name  : "",
        dataset_a: [],
        dataset_b: []
      }
    };
    if (!this.dataSet.rows) return finalResults;
    this.dataSet.rows.forEach(element => {
      let column = element.values[this.properties.label];
      if (!tempResults_a[column]){
        tempResults_a[column] = 0;
      }
      if (!tempResults_b[column]){
        tempResults_b[column] = 0;
      }
      tempResults_a[column]+=parseFloat(element.values[this.properties.value]);
      if (doubleChart) tempResults_b[column]+=parseFloat(element.values[this.properties.secondaryValue]);
    });

    for (let key in tempResults_a){
      finalResults.data.labels.push(key);
      finalResults.data.dataset_a.push(tempResults_a[key]);
    }
    if (doubleChart){
      for (let key in tempResults_b){
        finalResults.data.dataset_b.push(tempResults_b[key]);
      }
    }

    return finalResults;
  }

  private getValuesForKPICard(){
    let finalResults = {
      goal:0,
      value:0,
      difference:0,
      performance:0
    };
    
    if (!this.dataSet.rows) return finalResults;

    this.dataSet.rows.forEach(element => {
      finalResults.value+=parseFloat(element.values[parseInt(this.properties.value)].toString());
      finalResults.goal+=parseFloat(element.values[parseInt(this.properties.secondaryValue)].toString());
    });

    finalResults.difference = finalResults.value - finalResults.goal;
    finalResults.performance = finalResults.value/finalResults.goal;

    return finalResults;

  }

  private getValuesForPie(){

    let tempResults = {};
    if (!this.dataSet.rows) return [];
    this.dataSet.rows.forEach(element => {
      let column = element.values[this.properties.label];
      if (!tempResults[column]){
        tempResults[column] = 0;
      }
      tempResults[column]+=parseFloat(element.values[this.properties.value]);
    });

    let finalResults = {
      data: [],
      labels: []
    };
  
    let total = 0;
    for (let key in tempResults){
      total += tempResults[key];
      finalResults.data.push(tempResults[key]);
      finalResults.labels.push(key);
    }


    for (let i=0;i<finalResults.labels.length; i++){
      let relativeValue : number = Math.round(10000*parseFloat(finalResults.data[i])/total)/100;
      finalResults.labels[i]+=" ("+relativeValue+"%)";
    }

    return finalResults;
  }

  private getThemeColor(key:string){
    let color = undefined;
    if ((<any>window).__themeState__ && (<any>window).__themeState__.theme){
      let value = (<any>window).__themeState__.theme;
      if (value) return value[key];
    }
    return color;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
      
        {
          displayGroupsAsAccordion: false,
          groups: [
            {
              groupName : strings.DataSourceLabel,
              groupFields: [
                PropertyPaneDynamicField("bi_dataset", {
                  label:"",
                  propertyValueDepth: DynamicDataSharedDepth.None,
                  filters:{
                    propertyId:BiClientWebPart.DATA_KEY
                  }
                }),
                PropertyPaneTextField("title", {
                  label: strings.TitleLabel
                })
              ]
            },
            this.getWidgetGroup()
          ]
        }
      ]
    };
  }

  private getWidgetGroup():IPropertyPaneGroup{
    let groupFields: IPropertyPaneField<any>[]=[];
    let group : IPropertyPaneGroup = {
      groupName: strings.WidgetLabel,
      groupFields: groupFields
    };

    groupFields.push(PropertyPaneDropdown("widget", {
      label: strings.WidgetTypeLabel,
      options:[
        {key: WidgetType.Pie, text: strings.PieChartOption},
        {key: WidgetType.Chart, text: strings.ChartOption},
        {key: WidgetType.KPI, text: strings.KpiCardOption}
      ]
    }));

    if (this.properties.widget == WidgetType.Chart){
      groupFields.push(PropertyPaneDropdown("chartType", {
        label: strings.ChartType,
        options:[
          {key:ChartType.SingleBars, text: strings.BarsSingleOption},
          {key:ChartType.DoubleBars, text: strings.BarsDoubleOption},
          {key:ChartType.SingleLines, text: strings.LinesSingleOption},
          {key:ChartType.DoubleLines, text: strings.LinesSoubleOption},
        ]
      }));
    }

    if(this.properties.widget != WidgetType.KPI){
      groupFields.push(PropertyPaneDropdown("label", {
        label:strings.LabelsLabel,
        options:this.getColumnOptions([DataSetColumnType.Text])
      }));
    }

    groupFields.push(PropertyPaneDropdown("value", {
      label: strings.MainValueLabel,
      options:this.getColumnOptions([DataSetColumnType.Number, DataSetColumnType.Date])
    }));

    if (this.properties.widget == WidgetType.KPI || (this.properties.widget == WidgetType.Chart && (this.properties.chartType == ChartType.DoubleBars || this.properties.chartType == ChartType.DoubleLines))){
      groupFields.push( PropertyPaneDropdown("secondaryValue", {
        label: strings.SecondaryValueLabel,
        options:this.getColumnOptions([DataSetColumnType.Number, DataSetColumnType.Date]),
      }));
    }

    if (this.properties.widget != WidgetType.KPI){
      groupFields.push(PropertyPaneTextField("maxHeight", {
        label: strings.MaxHeightLabel,
        description : strings.MaxHeightDescription,
        deferredValidationTime: 1500,
        onGetErrorMessage: (value)=>{
          let errorMessage = strings.InvalidValueMessage;
          try{
            let size = parseInt (value);
            if (size >= 200){
              return "";
            } else {
              errorMessage = strings.MinValueMessage + BiClientWebPart.MIN_HEIGHT;
            }
          } catch (error){
            errorMessage = error;
          }
  
          return errorMessage;
        }
      }));
    }

    return group;

  }

  private getColumnOptions(types:DataSetColumnType[]) : IPropertyPaneDropdownOption[]{
    try{
    if (this.dataSet && this.dataSet.columns){
      let result = [];

        let columns : any = this.dataSet.columns;
        let index = 0;
        columns.forEach(element => {
          if(types.indexOf(element.type)!=-1) result.push({key: index.toString(), text:element.name});
          index++;
        });
        return result;
      
    }else{
      return  [];
    }
  }catch(error){
      return [];
    }
  }

}
