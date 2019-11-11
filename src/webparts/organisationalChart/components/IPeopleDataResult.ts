export interface IPeopleDataResult{
    RelevantResults:{
        TotalRows:number,
        Table:{
            Rows:[
                {
                    Key:string,
                    Value:string,
                    ValueType:string
                }
            ]
        };
    };
}