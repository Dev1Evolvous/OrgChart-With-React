export interface IDataNode {
    id: number;
    parent_id: number;
    name: string;
    children: Array<IDataNode>;
}

export class OrgChartNode implements IDataNode {
    public id: number;
    public parent_id: number;
    public name: string;
    public children: Array<IDataNode>;

    constructor(id: number, name: string, children?: Array<IDataNode>) {
        this.id = id;
        this.parent_id = id;
        this.name = name;
        this.children = children || [];
    }

    public addNode(node: IDataNode): void {
        this.children.push(node);
    }
}

