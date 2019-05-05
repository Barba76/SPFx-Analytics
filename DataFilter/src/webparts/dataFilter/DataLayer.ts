
import { SPHttpClient, SPHttpClientResponse, GraphHttpClient, GraphHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';


export class DataSet {
    public title : string;
    public description : string;
    public columns : Array<DataSetColumn>;
    public rows : Array<DataSetRow>;

    constructor() {
        this.title = "";
        this.description = "";
        this.columns = new Array<DataSetColumn>();
        this.rows = new Array<DataSetRow>();
    }
}

export enum DataSetColumnType {Text="text", Number="number", Date="date"}
export class DataSetColumn {
    public index : number;
    public name : string;
    public type : DataSetColumnType;
}

export class DataSetRow {
    public index: number;
    public values: Array<Object>;
}

export default class DataLayer {

    public static getMockData() : DataSet {
        
        let data : DataSet =  {
            title : "Mock Data",
            description : "Clubs Income",
            columns : [
                {index:0, name:"Year", type: DataSetColumnType.Text},
                {index:1, name:"Month", type: DataSetColumnType.Text},
                {index:2, name:"Type", type: DataSetColumnType.Text},
                {index:3, name:"Online Sales", type: DataSetColumnType.Number},
                {index:4, name:"Online Goals", type: DataSetColumnType.Number},
                {index:5, name:"Store Sales", type: DataSetColumnType.Number},
                {index:6, name:"Store Goals", type: DataSetColumnType.Number},
                {index:7, name:"Total Sales", type: DataSetColumnType.Number},
                {index:8, name:"Total Goals", type: DataSetColumnType.Number}
            ],
            rows : [
                {index:0, values:["2018","Jan","Online","102336","100000","0","0","102336","100000"]},
                {index:1, values:["2018","Jan","Store","0","0","27651","25000","27651","25000"]},
                {index:2, values:["2018","Feb","Online","139704","100000","0","0","139704","100000"]},
                {index:3, values:["2018","Feb","Store","0","0","17485","25000","17485","25000"]},
                {index:4, values:["2018","Mar","Online","79345","100000","0","0","79345","100000"]},
                {index:5, values:["2018","Mar","Store","0","0","19190","30000","19190","30000"]},
                {index:6, values:["2018","Apr","Online","137313","100000","0","0","137313","100000"]},
                {index:7, values:["2018","Apr","Store","0","0","28974","30000","28974","30000"]},
                {index:8, values:["2018","May","Online","82372","100000","0","0","82372","100000"]},
                {index:9, values:["2018","May","Store","0","0","17555","35000","17555","35000"]},
                {index:10, values:["2018","Jun","Online","80371","100000","0","0","80371","100000"]},
                {index:11, values:["2018","Jun","Store","0","0","41437","40000","41437","40000"]},
                {index:12, values:["2018","Jul","Online","134083","100000","0","0","134083","100000"]},
                {index:13, values:["2018","Jul","Store","0","0","56425","40000","56425","40000"]},
                {index:14, values:["2018","Aug","Online","88345","100000","0","0","88345","100000"]},
                {index:15, values:["2018","Aug","Store","0","0","29440","45000","29440","45000"]},
                {index:16, values:["2018","Sep","Online","148142","100000","0","0","148142","100000"]},
                {index:17, values:["2018","Sep","Store","0","0","33146","40000","33146","40000"]},
                {index:18, values:["2018","Oct","Online","85881","100000","0","0","85881","100000"]},
                {index:19, values:["2018","Oct","Store","0","0","33590","35000","33590","35000"]},
                {index:20, values:["2018","Nov","Online","66640","100000","0","0","66640","100000"]},
                {index:21, values:["2018","Nov","Store","0","0","31099","30000","31099","30000"]},
                {index:22, values:["2018","Dec","Online","59789","100000","0","0","59789","100000"]},
                {index:23, values:["2018","Dec","Store","0","0","28475","25000","28475","25000"]},
                {index:24, values:["2019","Jan","Online","85974","120000","0","0","85974","120000"]},
                {index:25, values:["2019","Jan","Store","0","0","35965","25000","35965","25000"]},
                {index:26, values:["2019","Feb","Online","157552","120000","0","0","157552","120000"]},
                {index:27, values:["2019","Feb","Store","0","0","36956","25000","36956","25000"]},
                {index:28, values:["2019","Mar","Online","135876","120000","0","0","135876","120000"]},
                {index:29, values:["2019","Mar","Store","0","0","27882","30000","27882","30000"]},
                {index:30, values:["2019","Apr","Online","175052","120000","0","0","175052","120000"]},
                {index:31, values:["2019","Apr","Store","0","0","25370","30000","25370","30000"]},
                {index:32, values:["2019","May","Online","117962","120000","0","0","117962","120000"]},
                {index:33, values:["2019","May","Store","0","0","36304","35000","36304","35000"]},
                {index:34, values:["2019","Jun","Online","87358","120000","0","0","87358","120000"]},
                {index:35, values:["2019","Jun","Store","0","0","21832","40000","21832","40000"]},
                {index:36, values:["2019","Jul","Online","161402","120000","0","0","161402","120000"]},
                {index:37, values:["2019","Jul","Store","0","0","57141","40000","57141","40000"]},
                {index:38, values:["2019","Aug","Online","138390","120000","0","0","138390","120000"]},
                {index:39, values:["2019","Aug","Store","0","0","31331","45000","31331","45000"]},
                {index:40, values:["2019","Sep","Online","66835","120000","0","0","66835","120000"]},
                {index:41, values:["2019","Sep","Store","0","0","46470","40000","46470","40000"]},
                {index:42, values:["2019","Oct","Online","112197","120000","0","0","112197","120000"]},
                {index:43, values:["2019","Oct","Store","0","0","44430","35000","44430","35000"]},
                {index:44, values:["2019","Nov","Online","100004","120000","0","0","100004","120000"]},
                {index:45, values:["2019","Nov","Store","0","0","30228","30000","30228","30000"]},
                {index:46, values:["2019","Dec","Online","80066","120000","0","0","80066","120000"]},
                {index:47, values:["2019","Dec","Store","0","0","15209","25000","15209","25000"]}
            ]
        };

        return data;

    }

    public static getDataFromSPList(context: WebPartContext, webUrl:string, id:string) : Promise<DataSet>{

        let site = webUrl || context.pageContext.web.absoluteUrl;
        return context.spHttpClient.get(`${site}/_api/web/lists/getById('${id}')/items?`, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {           

            return response.clone().json().then((json : any)=>{

                let dataSet = new DataSet();
                let rows = json.value;

                if (rows){
                    let rowIndex = 0;
                    rows.forEach(row => {
                        
                        let dataRow = new DataSetRow();
                        dataRow.index = rowIndex;
                        dataRow.values = new Array<string>();

                        let colIndex = 0;

                        for(var key in row){

                            let colVal = row[key];

                            if (rowIndex == 0){

                                let type : DataSetColumnType= DataSetColumnType.Text;
                                if (typeof colVal === "number"){
                                    type = DataSetColumnType.Number;
                                }/*else if (typeof colVal === "Date"){
                                    type = DataSetColumnType.Date;
                                }*/

                                let dataSetColumn = new DataSetColumn();
                                dataSetColumn.type = type;
                                dataSetColumn.name = key;
                                dataSetColumn.index = colIndex;

                                dataSet.columns.push(dataSetColumn);
                                colIndex ++;
                            }

                            dataRow.values.push(colVal? colVal.toString() : (colVal==0?"0":""));

                        }   
                        
                        dataSet.rows.push(dataRow);
                        rowIndex ++;
                    });

                }

                return Promise.resolve(dataSet);
            }).catch((error)=>{
                console.log(error);
                return Promise.resolve(new DataSet());
            });

        }).catch((error)=>{
            console.log(error);
            return Promise.resolve(new DataSet());
        });
        
    }

} 

