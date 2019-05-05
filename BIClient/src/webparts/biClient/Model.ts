export interface DataSet {
    title : string;
    description: string;
    columns : Array<DataSetColumn>;
    rows : Array<DataSetRow>;
}

export enum DataSetColumnType {Text="text", Number="number", Date="date"}
export interface DataSetColumn {
    index : number;
    name : string;
    type : DataSetColumnType;
}

export interface DataSetRow {
    index: number;
    values: Array<Object>;
}