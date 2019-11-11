export interface IOrgChartItem {
    Title: string;
    Guid: string;
    Id: number;
    parent_guid?: string;
    parent_id: number;
    Url?: string;
    Parents: any;
    ParentsId: number;
    JobTitle?: string;
    Manager?: string;
    Department?: string;
    EmpEmail?: string;
    Managers: any;
    Peers: any;
    Properties: any;

}
export interface IOrgChartItemGUID {
    Title: string;
    Url?: string;
    Guid: string;
    Parent_guid?: string;
    JobTitle?: string;
    Manager?: string;
    Department?: string;
    EmpEmail?: string;
    Peers: any;
    Properties: any;
}

export class ChartItem {
    public id: number;
    public title: string;
    public Url: string;
    public parent_id?: number;
    public JobTitle?: string;
    public Manager?: string;
    public Department?: string;
    public EmpEmail?: string;


    constructor(id: number, title: string, Url: string, parent_id?: number, JobTitle?: string,
        Manager?: string, Department?: string, EmpEmail?: string) {
        this.id = id;
        this.title = title;
        this.parent_id = parent_id;
        this.Url = Url;
        this.JobTitle = JobTitle;
        this.Manager = Manager;
        this.Department = Department;
        this.EmpEmail = EmpEmail;

    }
}

export class ChartItemGUID {

    public guid?: string;
    public title: string;
    public Url: string;
    public ID: number;
    public parent_guid: number;
    public JobTitle?: string;
    public Manager?: string;
    public Department?: string;
    public EmpEmail?: string;
    public Peers?: any;


    constructor(guid: string, title: string, Url: string, ID: number, parent_guid: number, JobTitle?: string,
        Manager?: string, Department?: string, EmpEmail?: string, Peers?: any) {
        this.guid = guid;
        this.title = title;
        this.parent_guid = parent_guid;
        this.Url = Url;
        this.JobTitle = JobTitle;
        this.Manager = Manager;
        this.Department = Department;
        this.EmpEmail = EmpEmail;
        this.Peers = Peers;
        this.ID = ID;

    }
}


export class HoldChartItemArrow {
    public nodeId: string;
    public nodeIdValue: string;

    constructor(nodeId: string, nodeIdValue: string) {
        this.nodeId = nodeId;
        this.nodeIdValue = nodeIdValue;
    }
}