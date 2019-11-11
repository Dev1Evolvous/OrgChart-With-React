import { Guid } from "@microsoft/sp-core-library";

export interface ITreeData {
   id?: number;
   ID?: number;
   ManagersId?: number;
   Title: string;
   Url: string;
   JobTitle?: string;
   Manager?: string;
   Department?: string;
   EmpEmail?: string;
   GUIDNew?: string;
  
}

export class TreeData {
  public id: number;
  public ID: number;
  public ManagersId: number;
  public Title: string;
  public Url: string;
  public JobTitle?: string;
  public Manager?: string;
  public Department?: string;
  public EmpEmail?: string;
  public GUIDNew?: string;

   constructor(id: number, ID: number, ManagersId: number, Title: string,
      Url: string, JobTitle?: string, Manager?: string, Department?: string,
      EmpEmail?: string, GUIDNew?: string) {
      this.id = id;
      this.ID = ID;
      this.Url = Url;
      this.Title = Title;
      this.ManagersId = ManagersId;
      this.JobTitle = JobTitle;
      this.Manager = Manager;
      this.Department = Department;
      this.EmpEmail = EmpEmail;
      this.GUIDNew=GUIDNew;
   }
}