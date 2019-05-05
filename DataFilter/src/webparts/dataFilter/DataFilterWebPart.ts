import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';

import styles from './DataFilterWebPart.module.scss';
import * as strings from 'DataFilterWebPartStrings';
import DataLayer, { DataSet, DataSetRow, DataSetColumnType } from './DataLayer';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface IDataFilterWebPartProps {
  mockData: boolean;
  listId:string;
  title: string;
  description: string;
  selectedColumns : string[];
  siteUrl : string;
  savePrefs : boolean;
  savedFilters :string [];
}

export default class DataFilterWebPart extends BaseClientSideWebPart<IDataFilterWebPartProps> implements IDynamicDataCallables  {

  private dataSet : DataSet;
  private allRows : Array<DataSetRow>;
  private filteredData: DataSet;
  private lastMockDataProperty : boolean;

  constructor(){
    super();
  }

  protected onInit(): Promise<void> {
        
    this.context.dynamicDataSourceManager.initializeSource(this);

    this.context.dynamicDataSourceManager.updateMetadata({
      title: this.properties.title,
      description : this.properties.description
    });
    

    this.lastMockDataProperty = this.properties.mockData;

    return this.retrieveData();
    
  }


  /**
   * Return list of dynamic data properties that this dynamic data source
   * returns
   */
  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'bi_dataset',
        title: 'Data Results'
      }
    ];
  }

  public _getDefaultAccessibleTitle() : any{
       return this.properties.description + "web part";
  }

  /**
   * Return the current value of the specified dynamic data set
   * @param propertyId ID of the dynamic data set to retrieve the value for
   */
  public getPropertyValue(propertyId: string): any {
    switch (propertyId) {
      case 'bi_dataset':
        return this.filteredData;

    }

    throw new Error('Bad property id');
  }

  public render(): void {

    if (this.renderedOnce && this.properties.mockData != this.lastMockDataProperty){
      this.lastMockDataProperty = this.properties.mockData;
      this.properties.selectedColumns = [];
      this.retrieveData().then(()=>{
        this.render();
      });
      return;
    }


    this.allRows = this.dataSet.rows;

    let html = `
      <div class="${ styles.dataFilter }">
        <div class="${ styles.container }">
          <div class="${styles.title}">
            <span>${this.properties.title}</span>
          </div>
          <div>`;
          if (this.dataSet && this.properties.selectedColumns){
            this.properties.selectedColumns.forEach(column=>{
              if (this.dataSet.columns[column]){
                let columName = this.dataSet.columns[column].name;
                html +=`<div class="${styles.group_items}">
                            <span class="${styles.item_title}" title="${columName}">${columName}</span>`;

                  this.getUniqueValuesFromSelectedColumn(column).forEach(element => {
                    html+=`      <span class="${styles.item}" data-selected="false" column="${column}" value="${element}">${element}</span>`;
                  });

                html += `</div>`;
              }
            });
          }
    html += `
          </div>
        </div>
      </div>`;

      this.domElement.innerHTML = html;

      let titles = this.domElement.getElementsByClassName(styles.item_title);
      for(let i=0; i< titles.length; i++){
        (<HTMLElement> titles[i]).onclick = ()=>{
          let parent = titles[i].parentElement;
          let childrenItems = parent.getElementsByClassName(styles.item);
          this.clearItemsSelection(childrenItems);
        };
      }

      let items = this.domElement.getElementsByClassName(styles.item);
      for(let i=0; i< items.length; i++){
        (<HTMLElement> items[i]).onclick = ()=>{this.toggleSelection(<HTMLElement> items[i]);};
      }

      if(this.properties.savePrefs){
            for (let i=0;i<items.length;i++){
              let itemValue = items[i].textContent;
              if (this.properties.savedFilters.indexOf(itemValue)!=-1){
                items[i].setAttribute("data-selected", "true");
                items[i].classList.add(styles.selected);
              }
            }
        }

      if(!this.renderedOnce ){
        this.updateFilteredData();
      } 

      this.lastMockDataProperty = this.properties.mockData;

  }

  private clearItemsSelection(items : any){

    for(let i=0; i<items.length;i++){
      items[i].setAttribute("data-selected","false");
      items[i].classList.remove(styles.selected);
    }
    this.updateFilteredData();
  }

  private getSelectedKeys() {
    let selectedKeys = [];
    if(!this.properties.selectedColumns) return selectedKeys;

    let items = this.domElement.getElementsByClassName(styles.item);
    for(let i=0; i< items.length; i++){
      if(items[i].getAttribute("data-selected") == "true"){
        selectedKeys.push(items[i].getAttribute("value"));
      }
    }

    this.properties.selectedColumns.forEach(column=>{
      let columnItems = [];
      let groupSelected = 0;
      for(let i=0; i< items.length; i++){
        if(items[i].getAttribute("column") == column) {
          columnItems.push(items[i]);
          if(items[i].getAttribute("data-selected") == "true"){
            groupSelected++;
            break;
          }
        }
      }

      if (groupSelected==0){
        columnItems.forEach(item=>{
          selectedKeys.push(item.getAttribute("value"));
        });
      }

    });

    return selectedKeys;
  }

  private updateFilteredData(){
    this.filteredData = this.getFilteredData();
    this.context.dynamicDataSourceManager.notifySourceChanged();
  }

  private getFilteredData():DataSet{
    let selectedKeys = this.getSelectedKeys();

    if (this.properties.savePrefs){
      this.properties.savedFilters = selectedKeys;
    } else {
      this.properties.savedFilters = [];
    }

    if (!selectedKeys.length) return this.dataSet;
    let filteredData = {
      title: this.properties.title,
      description: this.properties.description,
      columns: this.dataSet.columns,
      rows: []
    };
    let selectedCols = this.properties.selectedColumns;
    if (selectedCols){
      this.dataSet.rows.forEach(element => {
        let match=true; 
        for (var i=0; i < selectedCols.length; i++){
          let column = selectedCols[i];
          if (selectedKeys.some(value=>value == element.values[parseInt(column)]) == false) {
            match = false;
            break;
          }
        }
        if(match)filteredData.rows.push(element);
      });
    }
    return filteredData;
  }


  private toggleSelection(element:HTMLElement){
    if (element.getAttribute("data-selected") == "false"){
      element.setAttribute("data-selected", "true");
      element.classList.add(styles.selected);
    }else{
      element.setAttribute("data-selected", "false");
      element.classList.remove(styles.selected);
    }

    this.updateFilteredData();

  }


  private getUniqueValuesFromSelectedColumn(column) {
    let result = [];
    if (!this.allRows) return result;
    this.allRows.forEach(element => {
      let tempVal = element.values[parseInt(column)];
      if (result.some(value=> value == tempVal) == false){
        result.push(tempVal);
      }
    });
    return result;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private retrieveData () : Promise<any>{
    if (this.properties.mockData != true){
      return DataLayer.getDataFromSPList(this.context, this.properties.siteUrl, this.properties.listId).then((value : DataSet)=>{
        this.dataSet = value;
        this.context.propertyPane.refresh();
        Promise.resolve();
      });
    }else{
      this.dataSet = DataLayer.getMockData();
      return Promise.resolve();
    }
    
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.PropertiesLabel,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
              groupName: strings.DataSourceLabel,
              groupFields:[
                PropertyPaneToggle("mockData", {
                  label: "Mock Data"
                }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlLabel,
                  disabled: this.properties.mockData,
                  onGetErrorMessage: ()=>{return "";},
                  deferredValidationTime: 1000,
                }),
                PropertyFieldListPicker('listId', {
                  disabled: this.properties.mockData,
                  label: strings.SelectListLabel,
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.retrieveData.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  multiSelect: false,
                  webAbsoluteUrl: this.properties.siteUrl,
                  key: 'listPickerFieldId'
                }),
                PropertyFieldMultiSelect('selectedColumns', {
                  key: 'multiSelect',
                  label: strings.FilterFieldsLabel,
                  options:  this.dataToDropDown(),
                  selectedKeys : this.properties.selectedColumns || []
                }),
              ]
            },
            {
              groupName: strings.FilterOptionsLabel,
              groupFields:[
                PropertyPaneToggle("savePrefs", {
                  label: strings.SaveFiltersLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private dataToDropDown():IPropertyPaneDropdownOption[]{
    let data : IPropertyPaneDropdownOption[] = [];
    let object = this.dataSet;
    object.columns.forEach(element => {
      if (element.type == DataSetColumnType.Text){
        data.push({key:element.index, text: element.name});
      }
    });
    return data;
    
  }
}
