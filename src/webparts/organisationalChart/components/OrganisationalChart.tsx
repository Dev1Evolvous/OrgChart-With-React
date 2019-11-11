import * as React from 'react';
import styles from './OrganisationalChart.module.scss';
import { IOrganisationalChartProps } from './IOrganisationalChartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import OrgChart from 'react-orgchart';
import { IOrgChartViewerState } from './IOrgChartViewerState';
import { IOrgChartItem, ChartItem, HoldChartItemArrow, ChartItemGUID, IOrgChartItemGUID } from './IOrgChartItem';
import { IDataNode, OrgChartNode } from './OrgChartNode';
import { SPHttpClient, SPHttpClientResponse }
  from '@microsoft/sp-http';
import 'react-orgchart/index.css';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { initializeIcons } from '@uifabric/icons';
import { IPeopleDataResult } from './IPeopleDataResult';
import * as jquery from 'jquery';
import Draggable from 'react-draggable';

import { IconButton, BaseButton } from 'office-ui-fabric-react/lib/Button';
//import 'font-awesome/css/font-awesome.min.css';
import { SPComponentLoader } from '@microsoft/sp-loader';

import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import {
  CompactPeoplePicker,
  IBasePickerSuggestionsProps,
  IBasePicker,
  ListPeoplePicker,
  NormalPeoplePicker,
  ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { people, mru } from '@uifabric/example-data';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { assign } from 'office-ui-fabric-react/lib/Utilities';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import { Callout } from 'office-ui-fabric-react';
import { number } from 'prop-types';
import { UserProfileQuery } from '@pnp/sp/src/userprofiles';


const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading'
};

let countChilden: number = 1;
//let profileURL=this.props.pageContext.web.absoluteUrl;
let profileURL = "https://evolvous.sharepoint.com/";
export default class OrganisationalChart extends React.Component<IOrganisationalChartProps, any> {
  // All pickers extend from BasePicker specifying the item type.
  private _peopleList;
  private contextualMenuItems: IContextualMenuItem[] = [
    {
      key: 'newItem',
      icon: 'circlePlus',
      name: 'New'
    },
    {
      key: 'upload',
      icon: 'upload',
      name: 'Upload'
    },
    {
      key: 'divider_1',
      name: '-',
    },
    {
      key: 'rename',
      name: 'Rename'
    },
    {
      key: 'properties',
      name: 'Properties'
    },
    {
      key: 'disabled',
      name: 'Disabled item',
      disabled: true
    }
  ];


  public render(): React.ReactElement<IOrganisationalChartProps> {
    initializeIcons(undefined, { disableWarnings: true });
    let currentPicker: JSX.Element | undefined = undefined;
    currentPicker = this._renderLimitedSearch();
    //var email = this.props.pageContext.web.absoluteUrl;
    //email = this.props.pageContext.user.email;
    //alert(email);
    return (
      <div className={styles.organisationalChart}>
        <div className={styles.container}>
          <span className={styles.title} style={{ color: "#000000", padding: "5 0 0 15", textDecoration: "Bold" }}>{escape(this.props.description)}</span>
          <div className={styles.row}>
            <div className={styles.column}>

              <br />
              {/* <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a> */}

              <div className="ms-Grid" dir="ltr">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
                    <div style={{ width: "350px", float: "left" }} >
                      {currentPicker}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg12">
                    <br />
                    <Icon iconName="ReplyAll" className="ms-IconExample"
                      style={{ textAlign: "center", cursor: "Pointer", color: "#797bc9" }}
                      onClick={this.processUpNextManager} />
                    <Draggable
                      axis="x"
                      handle=".handle"
                      defaultPosition={{ x: 0, y: 0 }}
                      position={null}
                      grid={[25, 25]}
                      scale={1}
                      onDrag={this.handleDrag}>
                      <OrgChart style={{ cursor: "move" }} tree={this.state.orgChartItems}
                        NodeComponent={this.MyNodeComponent} />
                    </Draggable>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  /**
   * Takes in the picker input and modifies it in whichever way
   * the caller wants, i.e. parsing entries copied from Outlook (sample
   * input: "Aaron Reid <aaron>").
   *
   * @param input The text entered into the picker.
   */

  public constructor(props: IOrganisationalChartProps, any) {
    super(props);
    //SPComponentLoader.loadCss("font-awesome/css/font-awesome.min.css");
    //SPComponentLoader.loadCss("https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css");
    //SPComponentLoader.loadCss("https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css");
    // this.props.context.pageContext.user.loginName

    this.processUpNextManager = this.processUpNextManager.bind(this);
    this.processManagerNextDownUser = this.processManagerNextDownUser.bind(this);
    this.handleDrag = this.handleDrag.bind(this);

    this.state = {
      orgChartItems: [],
      LoadUpNextUser: "",
      currentPicker: 1,
      delayResults: false,
      selectedItems: [],
      personaState: [],
      textSearchUserValue: "",
      ManagerID: 0,
      ManagerNextDownUser: "",
      OrgChartNodesArray: [],
      DepthLevel: 0,
      StockDownUserId: [],
      StockDownUserIdValue: [],
      ManagerSetIdForExpand: 0,
      ManagerSetIdForExpandEmailID: 0,
      activeDrags: 0,
      deltaPosition: {
        x: 0, y: 0
      },
      controlledPosition: {
        x: -400, y: 200
      },
      StateChartItemGUIDNode: []
    };

    this.processOrgChartItems();
  }

  private MyNodeComponent = ({ node }) => {

    if (node.Id == 0) {
      return (
        <div className={styles.divparent}>
          <a href={"#"} className={styles.link}
            onClick={() => alert("No Data")} >
            {"No Data"}
          </a>
        </div>
      );
    }
    else if (node) {

      console.log(this.state.StockDownUserId);
      console.log(this.state.StockDownUserIdValue);
      let CollapseContentSingle = "StockDown";

      this.state.StockDownUserIdValue.forEach((elementRow) => {
        console.log(elementRow["nodeId"]);
        console.log(`#CollapseContentSingle${node.id}`);
        if (elementRow["nodeId"] === `#CollapseContentSingle${node.id}`) {

          CollapseContentSingle = elementRow["nodeIdValue"];
          return;
        }
      });

      return (
        <div className={styles.initechNode}>
          <table>
            <tr>
              <td className={styles.profileImg}>
                <img src={node.Url} ></img>
              </td>
            </tr>
            <tr>
              <td >
                <a href={"#"} className={styles.link}
                  onClick={() => alert(node.title)} >
                  <div style={{
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap", overflow: "hidden", width: "140px", color: "white"
                  }}> {node.title}
                  </div>
                </a>
              </td>
            </tr>
            <tr>
              <td className={styles.link} >
                <a href={"#"}
                  onClick={() => alert(node.EmpEmail)} >
                  <div style={{
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap", overflow: "hidden", width: "140px", color: "white"
                  }}> {node.EmpEmail}
                  </div>
                </a>
              </td>
            </tr>
            <tr>
              <td>
                <a href={"#"} className={styles.link}
                  onClick={() => alert(node.JobTitle)} >
                  <div style={{
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap", overflow: "hidden", width: "140px", color: "white"
                  }} > {node.JobTitle}
                  </div>
                </a>
              </td>
            </tr>
            <tr>
              <td>
                <Icon iconName={CollapseContentSingle.length > 0 ?
                  CollapseContentSingle : "StockDown"}
                  id={"#CollapseContentSingle" + node.id}
                  key={"CollapseContentSingle" + node.id} className="ms-IconExample"
                  style={{ textAlign: "center", cursor: "Pointer", color: "black" }}
                  onClick={() => this.processManagerNextDownUser(node.EmpEmail, node.id)} />
              </td>
            </tr>
          </table>
        </div>
      );
    }
    else {
      return (
        <div className={styles.initechNode}>"Up Next"</div>
      );
    }
  }

  public handleDrag = (e, ui) => {
    const { x, y } = this.state.deltaPosition;
    this.setState({
      deltaPosition: {
        x: x + ui.deltaX,
        y: y + ui.deltaY,
      }
    });
  }


  private userExists(EmpEmail) {
    //return this.state.OrgChartNodesArray.some((el) => {
    debugger;
    return this.state.StateChartItemGUIDNode.some((el) => {
      debugger;

      return el["EmpEmail"] === EmpEmail;
    });
  }

  private ElementIDExist(elementID) {
    return this.state.OrgChartNodesArray.some((el) => {
      return el["EmpEmail"] === elementID;
    });
  }

  private CountDepthLevel = (orgChartNodes) => {

    let countDepth = 0;
    let tempPrivousCount = 0;
    let tempLastCount = 0;
    let parent_id = 0;

    console.log(orgChartNodes);
    orgChartNodes.forEach((elementRow) => {

      if (elementRow["id"] === undefined || elementRow["id"] > 0) {
        tempPrivousCount = elementRow["id"];

        if (tempLastCount != tempPrivousCount && parent_id != elementRow["parent_id"]) {

          countDepth++;
          tempLastCount = tempPrivousCount;

          parent_id = elementRow["parent_id"];
        }
      }

    });
    this.setState({ DepthLevel: countDepth });
    console.log("Depth1: " + countDepth);
    console.log("Depth2: " + this.state.DepthLevel);


  }

  private pow = (x) => {
    let arrCount = [];

    if (x.length <= 0) {
      countChilden = Math.max.apply(null, arrCount);
      return countChilden;
    }
    else {

      console.log(x["children"]);
      if (x["children"] != undefined) {
        let innerArray = x["children"];
        let arr = [];
        // if (x["children"].length > 0) {
        innerArray.forEach((elementRow) => {
          try {

            if (elementRow["children"].length > 0) {
              var arrayToTree: any = require('array-to-tree');
              let orgChartHierarchyNodes: any;
              orgChartHierarchyNodes = arrayToTree(elementRow["children"]);

              let data = JSON.stringify(orgChartHierarchyNodes);


              let res1 = data.match(/children/g);

              if (res1.length > 0)
                arrCount.push(res1.length);
              else {
                arrCount.push(countChilden);
              }

              arr = elementRow["children"];
              return this.pow(arr);
            }
          }
          catch (err) {
          }
        });

        // let arr = x["children"];
        //countChilden++;
        //arrCount.push(2);
        return this.pow(arr);
        // }
      }
      else {
        try {
          let arr = [];

          let data = JSON.parse(x["children"]);


          var res = data.match(/children/g);

          if (res.length > 0)
            arrCount.push(res.length);
          else {
            arrCount.push(countChilden);
          }
          //countChilden++;
          return this.pow(arr);
        }

        catch (err) {
          arrCount.push(2);
        }
      }
    }
  }

  private convert = (orgChartNodes) => {

    var map = {};
    for (var i = 0; i < orgChartNodes.length; i++) {


      var obj = orgChartNodes[i];
      obj.children = [];

      map[obj.Id] = obj;

      var parent = obj.Parent || '-';
      if (!map[parent]) {
        map[parent] = {
          children: []
        };
      }
      map[parent].children.push(obj);
    }

    console.log("This is testing: " + map['-'].children);

  }

  private async processManagerNextDownUser(event, userId) {

    //alert("Hi" + event);
    await this.setState({
      ManagerNextDownUser: "i:0#.f|membership|" + event
    });
    this.readOrgChartItemsNextDownUserID().then((orgChartItemsOutter: IOrgChartItem[]) => {
      // logic will apply

      let peers = orgChartItemsOutter["DirectReports"];
      peers.forEach((elementRow) => {

        this.readOrgChartItemsDownUser(elementRow, userId)
          .then((orgChartItems: IOrgChartItem[]): void => {

            debugger;
            var result = [];
            jquery.each(this.state.StateChartItemGUIDNode, (i, e) => {
              var matchingItems = jquery.grep(result, (item) => {
                return item["EmpEmail"] === e["EmpEmail"];
              });
              if (matchingItems.length === 0) result.push(e);
            });

            if (result.length > 0) {
              this.setState({ StateChartItemGUIDNode: result });
            }



            debugger;
            let count = 0;
            let orgChartNodes: Array<ChartItem> = [];
            for (count = 0; count < this.state.StateChartItemGUIDNode.length; count++) {
              debugger;

              // this.userExists(this.state.StateChartItemGUIDNode[count]["EmpEmail"])
              orgChartNodes.push(new ChartItem(
                this.state.StateChartItemGUIDNode[count]["id"],
                this.state.StateChartItemGUIDNode[count]["title"],
                this.state.StateChartItemGUIDNode[count]["Url"],
                this.state.StateChartItemGUIDNode[count]["parent_id"],
                this.state.StateChartItemGUIDNode[count]["JobTitle"],
                "Manager",
                this.state.StateChartItemGUIDNode[count]["Department"],
                this.state.StateChartItemGUIDNode[count]["EmpEmail"]
              ));


              //let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
              // orgChartHoldNodesArrow.push(new HoldChartItemArrow(
              //   `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
            }

            debugger;
            console.log(orgChartNodes);

            //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
            //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
            this.setState({ OrgChartNodesArray: orgChartNodes });
            console.log("2: " + this.state.OrgChartNodesArray);
            //this.CountDepthLevel(orgChartNodes);

            var arrayToTree: any = require('array-to-tree');
            var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);


            var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
            this.setState({
              orgChartItems: JSON.parse(output)
            });

            debugger;

            //#region 
            // let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];

            // let orgChartNodes: Array<ChartItem> = [];
            // var count: number;
            // if (orgChartItems.length > 0) {

            //   for (count = 0; count < this.state.OrgChartNodesArray.length; count++) {

            //     orgChartNodes.push(new ChartItem(
            //       this.state.OrgChartNodesArray[count]["id"],
            //       this.state.OrgChartNodesArray[count]["title"],
            //       //this.state.OrgChartNodesArray[count].Url,
            //       "https://evolvous.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=accounts@evolvous.com",

            //       this.state.OrgChartNodesArray[count]["parent_id"] != undefined ?
            //         this.state.OrgChartNodesArray[count]["parent_id"] : undefined,

            //       this.state.OrgChartNodesArray[count]["JobTitle"],
            //       this.state.OrgChartNodesArray[count]["Manager"],
            //       this.state.OrgChartNodesArray[count]["Department"],
            //       this.state.OrgChartNodesArray[count]["EmpEmail"]
            //     ));

            //     // orgChartHoldNodesArrow.push(new HoldChartItemArrow(
            //     //   `CollapseContentSingle${this.state.OrgChartNodesArray[count].id}`,
            //     //   "StockDown"));

            //     //this.setState({ StockDownUserId: `CollapseContentSingle${userId}` });
            //     //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
            //   }
            //   count = 0;
            //   for (count = 0; count < orgChartItems.length; count++) {

            //     // orgChartItems.forEach((row) => {
            //     //   orgChartNodes.push(new ChartItem(
            //     //     row.Id,
            //     //     row["Title"],
            //     //     row.Url,
            //     //     row.Managers ? row.Managers.ID : undefined,
            //     //     row["JobTitle"],
            //     //     row["Manager"],
            //     //     row["Department"],
            //     //     row["EmpEmail"]));
            //     // });

            //     // Check already exist of not
            //     // this.state.OrgChartNodesArray.some((el) => {

            //     //   return el["EmpEmail"] === orgChartItems[count]["EmpEmail"];
            //     // });


            //     if (!this.userExists(orgChartItems[count]["EmpEmail"])) {
            //       orgChartNodes.push(new ChartItem(
            //         orgChartItems[count].Id,
            //         orgChartItems[count]["Title"],
            //         "https://evolvous.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=accounts@evolvous.com",

            //         //orgChartItems[count].Url,
            //         orgChartItems[count].Managers ? orgChartItems[count].Managers.ID : undefined,
            //         orgChartItems[count]["JobTitle"],
            //         orgChartItems[count]["Manager"],
            //         orgChartItems[count]["Department"],
            //         orgChartItems[count]["EmpEmail"]
            //       ));



            //       if (this.state.ManagerSetIdForExpand === orgChartItems[count].Id)
            //         orgChartHoldNodesArrow.push(new HoldChartItemArrow(
            //           `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
            //       // else {
            //       //   orgChartHoldNodesArrow.push(new HoldChartItemArrow(
            //       //     `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
            //       // }
            //       //this.setState({ StockDownUserId: `#StockDown${userId}` });
            //       //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });

            //     }
            //     else {
            //       orgChartNodes = orgChartNodes.filter((item) =>
            //         item.EmpEmail !== orgChartItems[count]["EmpEmail"]
            //       );


            //       if (this.state.ManagerSetIdForExpand === orgChartItems[count].Id)
            //         orgChartHoldNodesArrow.push(new HoldChartItemArrow(
            //           `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
            //       //#region 
            //       // else {
            //       //   orgChartHoldNodesArrow.push(new HoldChartItemArrow(
            //       //     `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
            //       // }


            //       //this.setState({ StockDownUserId: `#StockDown${userId}` });
            //       //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
            //       //#endregion
            //     }
            //   }
            //   //#region
            //   // else {

            //   //   orgChartNodes.push(new ChartItem(
            //   //     orgChartItemsOutter[0].Id,
            //   //     orgChartItemsOutter[0]["Title"],
            //   //     null,
            //   //     undefined,
            //   //     orgChartItemsOutter[0]["JobTitle"],
            //   //     orgChartItemsOutter[0]["Manager"],
            //   //     orgChartItemsOutter[0]["Department"],
            //   //     orgChartItemsOutter[0]["EmpEmail"]));

            //   //   for (count = 0; count < orgChartItems.length; count++) {

            //   //     orgChartNodes.push(new ChartItem(
            //   //       orgChartItems[count].Id,
            //   //       orgChartItems[count]["Title"],
            //   //       orgChartItems[count].Url,
            //   //       orgChartItems[count].Managers ? orgChartItems[count].Managers.ID : undefined,
            //   //       orgChartItems[count]["JobTitle"],
            //   //       orgChartItems[count]["Manager"],
            //   //       orgChartItems[count]["Department"],
            //   //       orgChartItems[count]["EmpEmail"]
            //   //     ));
            //   //   }
            //   // }
            //   //#endregion
            // }

            // let orgChartHoldNodesArrowTemp: Array<HoldChartItemArrow> = [];
            // orgChartHoldNodesArrowTemp = this.state.StockDownUserIdValue;

            // if (orgChartHoldNodesArrow.length > 0) {
            //   //#region 
            //   // if (this.userExists(`#CollapseContentSingle${orgChartHoldNodesArrow[0].nodeId}`)) {
            //   //   orgChartNodes = orgChartNodes.filter((item) =>
            //   //       item.EmpEmail !== orgChartItems[count]["EmpEmail"]
            //   //     );
            //   // }
            //   //#endregion

            //   orgChartHoldNodesArrowTemp.push(new HoldChartItemArrow(
            //     `#CollapseContentSingle${orgChartHoldNodesArrow[0].nodeId}`, "StockDown"));
            // }
            // else {

            //   orgChartHoldNodesArrowTemp.push(new HoldChartItemArrow(
            //     `#CollapseContentSingle${this.state.ManagerSetIdForExpand}`, "StockDown"));
            // }

            // this.setState({ StockDownUserIdValue: orgChartHoldNodesArrowTemp });

            // //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });


            // var arrayToTree: any = require('array-to-tree');
            // let orgChartHierarchyNodes: any;
            // if (orgChartNodes.length > 0) {
            //   this.setState({ OrgChartNodesArray: orgChartNodes });
            //   orgChartHierarchyNodes = arrayToTree(orgChartNodes);
            //   if (this.state.OrgChartNodesArray.length > 0)
            //     this.setState({ LoadUpNextUser: this.state.OrgChartNodesArray[0]["EmpEmail"] });
            //   else {
            //     this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
            //   }

            // }
            // else {
            //   orgChartHierarchyNodes = arrayToTree(this.state.OrgChartNodesArray);
            //   if (this.state.OrgChartNodesArray.length > 0)
            //     this.setState({ LoadUpNextUser: this.state.OrgChartNodesArray[0]["EmpEmail"] });
            //   else {
            //     this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
            //   }

            // }

            // //var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);
            // //this.convert(orgChartHierarchyNodes);

            // var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
            // this.pow(orgChartHierarchyNodes[0]);

            // if (output.length > 0)
            //   this.setState({
            //     orgChartItems: JSON.parse(output)
            //   });
            // //this.setState({ loadUpNext: false });
            //#endregion

          });
      });
    });
  }

  private async  readOrgChartItemsNextDownUserID(): Promise<IOrgChartItem[]> {

    let url = "";
    if (this.state.ManagerNextDownUser.length > 0) {
      //   url =
      //     `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?$top 2&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID
      // &$orderby=Title asc&$filter=startswith(EmpEmail,'${this.state.ManagerNextDownUser}')`;

      url = "" + this.props.pageContext.web.absoluteUrl
        + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
        + encodeURIComponent('' + this.state.ManagerNextDownUser + '') + "'";
    }
    else {
      //CurrentUser will Callout;
      // url = `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID
      //   &$orderby=Title asc&$filter=startswith(EmpEmail,'${this.props.pageContext.user.email}')`;

      url = "" + this.props.pageContext.web.absoluteUrl
        + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
        + encodeURIComponent('' + this.props.pageContext.user.email + '') + "'";
    }

    return await new Promise<IOrgChartItem[]>(
      (resolve, reject) =>
        //fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
        fetch(url,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response) => {
            return response.json();
          })
          .then((response): void => {


            // this.setState({ ManagerID: managerID });
            // console.log(this.state.ManagerID)

            resolve(response);
          }, (error: any): void => {
            reject(0);
          }));
  }

  private async readOrgChartItemsDownUser(managerEmailID, userID): Promise<IOrgChartItem[]> {

    let url = "" + this.props.pageContext.web.absoluteUrl
      + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
      + encodeURIComponent('' + managerEmailID + '') + "'";

    let managerEmailID1 = "charan.sodhi@evolvous.com";
    return await new Promise<IOrgChartItem[]>(
      (resolve, reject) =>
        //fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
        fetch(url,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).

          // }).then((response: SPHttpClientResponse): Promise<{ value: IOrgChartItem[] }> => {
          //   
          //   return response.json();
          // })
          // .then((response: { value: IOrgChartItem[] }): void => {
          //   
          then((response) => {

            return response.json();
          })
          .then((response): void => {


            //var userDisplayName = response.d.DisplayName;
            //var AccountName = response.d.AccountName;
            let orgChartNodes: Array<ChartItem> = [];
            let imgURL: string;
            let userName: string;
            let guid: string;
            let parent_guid: number;
            let jobTitle: string;
            let emailID: string;
            let managers: string;
            let department: string;
            let peers: any;
            let ID: number;
            let properties = response.UserProfileProperties;
            peers = response.Peers;
            let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
            let ChartItemGUIDNode: Array<ChartItemGUID> = [];
            let ChartItemGUIDNode1: Array<IOrgChartItem> = [];
            var count: number;

            orgChartNodes = this.state.StateChartItemGUIDNode;

            for (var i = 0; i < properties.length; i++) {

              var property = properties[i];
              if (property.Key == "WorkEmail") {
                emailID = property.Value;
              }
              if (property.Key == "WorkEmail") {
                imgURL = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${property.Value}`;
              }
              if (property.Key == "PreferredName") {
                userName = response.DisplayName;
              }
              if (property.Key == "UserProfile_GUID") {
                guid = property.Value;
              }
              if (property.Key == "SPS-JobTitle") {
                jobTitle = property.Value;
              }
              if (property.Key == "Department") {
                department = property.Value;
              }
              if (property.Key == "SPS-SharePointHomeExperienceState") {
                ID = property.Value;
              }
            }

            parent_guid = userID;

            orgChartNodes.push(new ChartItem(
              ID,
              userName,
              imgURL,
              parent_guid,
              jobTitle,
              "Managers",
              department,
              emailID)
            );

            //this.setState({ StateChartItemGUIDNode: ChartItemGUIDNode });
            this.setState({ StateChartItemGUIDNode: orgChartNodes });

            var result = [];
            jquery.each(this.state.StateChartItemGUIDNode, (ii, e) => {
              var matchingItems = jquery.grep(result, (item) => {
                return item["EmpEmail"] === e["EmpEmail"];
              });
              if (matchingItems.length === 0) result.push(e);
            });

            if (result.length > 0) {
              this.setState({ StateChartItemGUIDNode: result });
            }
            console.log(this.state.StateChartItemGUIDNode);

            //ChartItemGUIDNode1[0].Peers = response.Peers;

            //this.setState({ ManagerSetIdForExpand: managerid });
            // this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
            // //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
            // this.setState({ OrgChartNodesArray: orgChartNodes });
            // console.log("2: " + this.state.OrgChartNodesArray);
            // //this.CountDepthLevel(orgChartNodes);
            // var arrayToTree: any = require('array-to-tree');
            // var orgChartHierarchyNodes: any = arrayToTree(this.state.StateChartItemGUIDNode);
            // var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
            // this.setState({
            //   orgChartItems: JSON.parse(output)
            // });
            debugger;
            let relevantResults: any = response.value;
            resolve(response);
          }, (error: any): void => {
            reject(0);
          }));
  }

  private async processUpNextManager() {
    // this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["Title"] })
    //alert(this.state.LoadUpNextUser);

    this.readOrgChartItemsLoadUpNextUser().then((orgChartItemsOutter: IOrgChartItem[]) => {


      console.log(orgChartItemsOutter["ExtendedManagers"]);
      let temp = orgChartItemsOutter["ExtendedManagers"];

      let temp1 = `i:0#.f|membership|${this.state.StateChartItemGUIDNode[0]["EmpEmail"]}`;

      if (orgChartItemsOutter["ExtendedManagers"] !== temp1) {


        let managerArray = orgChartItemsOutter["UserProfileProperties"];
        let userManagerID = managerArray["Manager"];
        for (var ii = 0; ii < managerArray.length; ii++) {

          var property = managerArray[ii];
          if (property.Key == "Manager") {
            userManagerID = property.Value;
          }
        }
        debugger;

        this.readOrgChartItemsMiddleUpNext(userManagerID)
          .then((orgChartItemsMiddle: IOrgChartItem[]) => {

            let peers = orgChartItemsMiddle["DirectReports"];

            peers.forEach((elementRow) => {

              this.readOrgChartItems(elementRow)
                .then((orgChartItems: IOrgChartItem[]): void => {
                  let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
                  let orgChartNodes: Array<ChartItem> = [];
                  var count: number;

                  var result = [];
                  jquery.each(this.state.StateChartItemGUIDNode, (i, e) => {
                    var matchingItems = jquery.grep(result, (item) => {
                      return item["EmpEmail"] === e["EmpEmail"];
                    });
                    if (matchingItems.length === 0) result.push(e);
                  });

                  if (result.length > 0) {
                    this.setState({ StateChartItemGUIDNode: result });
                  }

                  for (count = 0; count < this.state.StateChartItemGUIDNode.length; count++) {


                    orgChartNodes.push(new ChartItem(
                      this.state.StateChartItemGUIDNode[count]["id"],
                      this.state.StateChartItemGUIDNode[count]["title"],
                      this.state.StateChartItemGUIDNode[count]["Url"],
                      this.state.StateChartItemGUIDNode[count]["parent_id"],
                      this.state.StateChartItemGUIDNode[count]["JobTitle"],
                      this.state.StateChartItemGUIDNode[count]["Manager"],
                      this.state.StateChartItemGUIDNode[count]["Department"],
                      this.state.StateChartItemGUIDNode[count]["EmpEmail"]
                    ));


                    //let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
                    // orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                    //   `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
                  }
                  debugger;
                  this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
                  //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
                  this.setState({ OrgChartNodesArray: orgChartNodes });
                  console.log("2: " + this.state.OrgChartNodesArray);
                  //this.CountDepthLevel(orgChartNodes);

                  var arrayToTree: any = require('array-to-tree');
                  debugger;
                  var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);
                  debugger;

                  var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
                  this.setState({
                    orgChartItems: JSON.parse(output)
                  });
                  debugger;



                  //#region 
                  // let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
                  // let orgChartNodes: Array<ChartItem> = [];
                  // var count: number;

                  // if (userManagerID != undefined) {
                  //   if (orgChartItems.length > 0) {
                  //     orgChartNodes.push(new ChartItem(
                  //       orgChartItemsMiddle[0].Id,
                  //       orgChartItemsMiddle[0]["Title"],
                  //       "https://evolvous.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=accounts@evolvous.com",

                  //       //null,
                  //       undefined,
                  //       orgChartItemsMiddle[0]["JobTitle"],
                  //       orgChartItemsMiddle[0]["Manager"],
                  //       orgChartItemsMiddle[0]["Department"],
                  //       orgChartItemsMiddle[0]["EmpEmail"]));


                  //     if (orgChartItems.length > 1)
                  //       orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                  //         `#CollapseContentSingle${orgChartItemsMiddle[0].Id}`, "StockDown"));
                  //     else {
                  //       orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                  //         `#CollapseContentSingle${orgChartItemsMiddle[0].Id}`, "StockDown"));
                  //     }

                  //     //this.setState({ StockDownUserId: `#StockDown${userId}` });

                  //   }
                  // }
                  // else {
                  //   if (orgChartItemsMiddle.length > 0) {
                  //     orgChartNodes.push(new ChartItem(
                  //       orgChartItemsMiddle[0].Id,
                  //       orgChartItemsMiddle[0]["Title"],
                  //       "https://evolvous.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=accounts@evolvous.com",

                  //       //null,
                  //       undefined,
                  //       orgChartItemsMiddle[0]["JobTitle"],
                  //       orgChartItemsMiddle[0]["Manager"],
                  //       orgChartItemsMiddle[0]["Department"],
                  //       orgChartItemsMiddle[0]["EmpEmail"]));

                  //     // orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                  //     //   `#CollapseContentSingle${orgChartItemsMiddle[0].Id}`, "StockDown"));

                  //     if (orgChartItems.length > 1)
                  //       orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                  //         `#CollapseContentSingle${orgChartItemsMiddle[0].Id}`, "StockDown"));
                  //     else {
                  //       orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                  //         `#CollapseContentSingle${orgChartItemsMiddle[0].Id}`, "StockDown"));
                  //     }

                  //     //this.setState({ StockDownUserId: `#StockDown${userId}` });
                  //     //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
                  //   }
                  // }
                  // for (count = 0; count < orgChartItems.length; count++) {

                  //   orgChartNodes.push(new ChartItem(
                  //     orgChartItems[count].Id,
                  //     orgChartItems[count]["Title"],
                  //     "https://evolvous.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=accounts@evolvous.com",

                  //     //orgChartItems[count].Url,
                  //     orgChartItems[count].Managers ? orgChartItems[count].Managers.ID
                  //       : undefined,
                  //     orgChartItems[count].JobTitle,
                  //     orgChartItems[count].Manager,
                  //     orgChartItems[count].Department,
                  //     orgChartItems[count].EmpEmail));

                  //   //let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
                  //   orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                  //     `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));

                  //   //this.setState({ StockDownUserId: `#StockDown${userId}` });
                  //   //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
                  // }
                  // this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
                  // this.setState({ LoadUpNextUser: orgChartItemsMiddle[0]["EmpEmail"] });
                  // this.setState({ OrgChartNodesArray: orgChartNodes });

                  // //this.CountDepthLevel(orgChartNodes);

                  // // for (count = 0; count < this.state.OrgChartNodesArray.length; count++) {

                  // //   console.log(`1: ${count} - ${this.state.OrgChartNodesArray[count].title}`);
                  // // }
                  // // }

                  // var arrayToTree: any = require('array-to-tree');
                  // var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);
                  // var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
                  // this.setState({ orgChartItems: JSON.parse(output) });
                  // this.pow(output);
                  // //this.setState({ loadUpNext: false });
                  //#endregion
                });
            });
          });
      }
    });

  }

  private async readOrgChartItemsLoadUpNextUser(): Promise<IOrgChartItem[]> {

    let url =
      `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID
    &$orderby=Title asc&$filter=startswith(EmpEmail,'${this.state.LoadUpNextUser}')`;

    url = "" + this.props.pageContext.web.absoluteUrl
      + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
      + encodeURIComponent('i:0#.f|membership|' + this.state.StateChartItemGUIDNode[0]["EmpEmail"] + '') + "'";

    return await new Promise<IOrgChartItem[]>(
      (resolve, reject) =>
        //fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
        fetch(url,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response) => {
            return response.json();
          })
          .then((response): void => {

            debugger;

            // this.setState({ ManagerID: managerID });
            // console.log(this.state.ManagerID)

            resolve(response);
          }, (error: any): void => {
            reject(0);
          }));
  }

  private async readOrgChartItemsMiddleUpNext(managerAsUserid): Promise<IOrgChartItem[]> {

    let url = "";
    // if (managerAsUserid === undefined) {
    //   url = `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID&$orderby=Title asc&$filter=ID eq 1`;
    // }
    // else {
    //   url = `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID&$orderby=Title asc&$filter=ID eq ${managerAsUserid}`;
    // }
    debugger;
    url = "" + this.props.pageContext.web.absoluteUrl
      + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
      + encodeURIComponent('' + managerAsUserid + '') + "'";

    return await new Promise<IOrgChartItem[]>(
      (resolve: (itemId: IOrgChartItem[]) => void, reject: (error: any) => void): void => {
        this.props.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response) => {
            return response.json();
          })
          .then((response): void => {
            debugger;

            // this.setState({ ManagerID: managerID });
            // console.log(this.state.ManagerID)
            let orgChartNodes: Array<ChartItem> = [];
            let imgURL: string;
            let userName: string;
            let guid: string;
            let parent_guid: number;
            let jobTitle: string;
            let emailID: string;
            let managers: string;
            let department: string;
            let peers: any;
            let ID: number;
            let properties = response.UserProfileProperties;
            peers = response.Peers;
            let ChartItemGUIDNode: Array<ChartItemGUID> = [];

            for (var i = 0; i < properties.length; i++) {
              var property = properties[i];
              if (property.Key == "WorkEmail") {
                emailID = property.Value;
              }
              if (property.Key == "WorkEmail") {
                imgURL = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${property.Value}`;
              }
              if (property.Key == "PreferredName") {
                userName = response.DisplayName;
              }
              if (property.Key == "UserProfile_GUID") {
                guid = property.Value;
              }
              if (property.Key == "SPS-JobTitle") {
                jobTitle = property.Value;
              }
              if (property.Key == "Department") {
                department = property.Value;
              }
              if (property.Key == "SPS-SharePointHomeExperienceState") {
                ID = property.Value;
              }
            }

            parent_guid = undefined;

            orgChartNodes.push(new ChartItem(
              ID,
              userName,
              imgURL,
              parent_guid,
              jobTitle,
              "managers",
              department,
              emailID)
            );

            this.setState({ StateChartItemGUIDNode: orgChartNodes });

            console.log(this.state.StateChartItemGUIDNode);
            let relevantResults: any = response.value;
            resolve(response);
          }, (error: any): void => {
            reject(error);
          });
      });
  }

  private async readOrgChartItemsID(): Promise<IOrgChartItem[]> {

    let url = "";
    // if (this.state.textSearchUserValue.length > 0)
    //   url =
    //     `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID&$orderby=Title asc&$filter=startswith(EmpEmail,'${this.state.textSearchUserValue}')`;
    // else {
    //   //CurrentUser will Callout;
    //   url = `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID&$orderby=Title asc&$filter=startswith(EmpEmail,'${this.props.pageContext.user.email}')`;
    // }

    if (this.state.textSearchUserValue.length > 0) {
      url =
        "" + this.props.pageContext.web.absoluteUrl
        + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
        + encodeURIComponent('i:0#.f|membership|' + this.state.textSearchUserValue + '') + "'";
    }
    else {
      //CurrentUser will Callout;
      url = "" + this.props.pageContext.web.absoluteUrl
        + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
        + encodeURIComponent('i:0#.f|membership|' + this.props.pageContext.user.email + '') + "'";

      // url = "" + this.props.pageContext.web.absoluteUrl
      //   + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
      //   + encodeURIComponent('i:0#.f|membership|charan.sodhi@evolvous.com') + "'";
    }

    return await new Promise<IOrgChartItem[]>(
      (resolve, reject) =>
        //fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
        fetch(url,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response) => {
            return response.json();
          })
          .then((response): void => {

            let relevantResults: any = response.value;
            // this.setState({ ManagerID: managerID });
            // console.log(this.state.ManagerID)
            let orgChartNodes: Array<ChartItem> = [];
            let imgURL: string;
            let userName: string;
            let guid: string;
            let parent_guid: number;
            let jobTitle: string;
            let emailID: string;
            let managers: string;
            let department: string;
            let peers: any;
            let ID: number;
            let properties = response.UserProfileProperties;
            peers = response.Peers;
            let ChartItemGUIDNode: Array<ChartItemGUID> = [];

            for (var i = 0; i < properties.length; i++) {
              var property = properties[i];
              if (property.Key == "WorkEmail") {
                emailID = property.Value;
              }
              if (property.Key == "WorkEmail") {
                imgURL = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${property.Value}`;
              }
              if (property.Key == "PreferredName") {
                userName = response.DisplayName;
              }
              if (property.Key == "UserProfile_GUID") {
                guid = property.Value;
              }
              if (property.Key == "SPS-JobTitle") {
                jobTitle = property.Value;
              }
              if (property.Key == "Department") {
                department = property.Value;
              }
              if (property.Key == "SPS-SharePointHomeExperienceState") {
                ID = property.Value;
              }
            }

            parent_guid = undefined;

            orgChartNodes.push(new ChartItem(
              ID,
              userName,
              imgURL,
              parent_guid,
              jobTitle,
              "managers",
              department,
              emailID)
            );

            this.setState({ StateChartItemGUIDNode: orgChartNodes });

            console.log(this.state.StateChartItemGUIDNode);

            resolve(response);
          }, (error: any): void => {
            reject(0);
          }));
  }

  private async readOrgChartItems(managerEmailID): Promise<IOrgChartItem[]> {

    let url = "" + this.props.pageContext.web.absoluteUrl
      + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
      + encodeURIComponent('' + managerEmailID + '') + "'";

    let managerEmailID1 = "charan.sodhi@evolvous.com";
    return await new Promise<IOrgChartItem[]>(
      (resolve, reject) =>
        //fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
        fetch(url,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).

          // }).then((response: SPHttpClientResponse): Promise<{ value: IOrgChartItem[] }> => {
          //   
          //   return response.json();
          // })
          // .then((response: { value: IOrgChartItem[] }): void => {
          //   
          then((response) => {

            return response.json();
          })
          .then((response): void => {


            //var userDisplayName = response.d.DisplayName;
            //var AccountName = response.d.AccountName;
            let orgChartNodes: Array<ChartItem> = [];
            let imgURL: string;
            let userName: string;
            let guid: string;
            let parent_guid: number;
            let jobTitle: string;
            let emailID: string;
            let managers: string;
            let department: string;
            let peers: any;
            let ID: number;

            let properties = response.UserProfileProperties;
            peers = response.Peers;
            let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
            let ChartItemGUIDNode: Array<ChartItemGUID> = [];
            let ChartItemGUIDNode1: Array<IOrgChartItem> = [];
            var count: number;

            orgChartNodes = this.state.StateChartItemGUIDNode;

            for (var i = 0; i < properties.length; i++) {

              var property = properties[i];
              if (property.Key == "WorkEmail") {
                emailID = property.Value;
              }
              if (property.Key == "WorkEmail") {
                imgURL = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${property.Value}`;
              }
              if (property.Key == "PreferredName") {
                userName = response.DisplayName;
              }
              if (property.Key == "UserProfile_GUID") {
                guid = property.Value;
              }
              if (property.Key == "SPS-JobTitle") {
                jobTitle = property.Value;
              }
              if (property.Key == "Department") {
                department = property.Value;
              }
              if (property.Key == "SPS-SharePointHomeExperienceState") {
                ID = property.Value;
              }

            }


            debugger;
            parent_guid = this.state.StateChartItemGUIDNode[0].id;


            orgChartNodes.push(new ChartItem(
              ID,
              userName,
              imgURL,
              parent_guid,
              jobTitle,
              "managers",
              department,
              emailID)
            );


            //this.setState({ StateChartItemGUIDNode: ChartItemGUIDNode });
            this.setState({ StateChartItemGUIDNode: orgChartNodes });

            console.log(this.state.StateChartItemGUIDNode);

            //ChartItemGUIDNode1[0].Peers = response.Peers;

            //this.setState({ ManagerSetIdForExpand: managerid });
            // this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
            // //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
            // this.setState({ OrgChartNodesArray: orgChartNodes });
            // console.log("2: " + this.state.OrgChartNodesArray);
            // //this.CountDepthLevel(orgChartNodes);
            // var arrayToTree: any = require('array-to-tree');
            // var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);
            // var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
            // this.setState({
            //   orgChartItems: JSON.parse(output)
            // });
            let relevantResults: any = response.value;
            resolve(response);
          }, (error: any): void => {
            reject(0);
          }));
  }

  private async readOrgChartItemsTemp(managerid): Promise<IOrgChartItem[]> {

    let managerEmailID1 = "charan.sodhi@evolvous.com";
    return await new Promise<IOrgChartItem[]>(
      (resolve: (itemId: IOrgChartItem[]) => void, reject: (error: any) => void): void => {
        this.props.spHttpClient.get("" + this.props.pageContext.web.absoluteUrl
          + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
          + encodeURIComponent('i:0#.f|membership|' + managerEmailID1 + '') + "'",
          // this.props.spHttpClient.get(
          //   `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?$top 2&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID&$orderby=Title asc&$filter=Managers/ID eq ${managerid}`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
            // }).then((response: SPHttpClientResponse): Promise<{ value: IOrgChartItem[] }> => {
            //   
            //   return response.json();
            // })
            // .then((response: { value: IOrgChartItem[] }): void => {
            //   
          }).then((response) => {

            return response.json();
          })
          .then((response): void => {


            //var userDisplayName = response.d.DisplayName;
            //var AccountName = response.d.AccountName;
            let orgChartNodes: Array<ChartItem> = [];
            let imgURL: string;
            let userName: string;
            let guid: string;
            let parent_guid: number;
            let jobTitle: string;
            let emailID: string;
            let managers: string;
            let department: string;


            var properties = response.UserProfileProperties;
            let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
            let ChartItemGUIDNode: Array<ChartItemGUID> = [];
            let ChartItemGUIDNode1: Array<IOrgChartItem> = [];
            var count: number;

            for (var i = 0; i < properties.length; i++) {
              var property = properties[i];
              if (property.Key == "WorkEmail") {
                emailID = property.Value;
              }
              if (property.Key == "WorkEmail") {
                imgURL = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${property.Value}`;
              }
              if (property.Key == "PreferredName") {
                userName = response.DisplayName;
              }
              if (property.Key == "UserProfile_GUID") {

                guid = property.Value;
              }
              if (property.Key == "SPS-JobTitle") {
                jobTitle = property.Value;
              }
              if (property.Key == "Department") {
                department = property.Value;
              }
            }


            parent_guid = undefined;

            // ChartItemGUIDNode.push(new ChartItemGUID(
            //   guid,
            //   userName,
            //   imgURL,
            //   parent_guid,
            //   jobTitle,
            //   managers,
            //   department,
            //   emailID)
            // );

            this.setState({ StateChartItemGUIDNode: ChartItemGUIDNode });

            // if (orgChartItemsOutter.length > 0) { }

            let peers = response.Peers;


            peers.forEach((elementEmail) => {


              this.readOrgChartItemsArray(elementEmail)
                .then((responseData) => {
                  return JSON.stringify(responseData);
                }).then((responseData) => {


                  let properties1 = responseData;
                  //properties = response.UserProfileProperties;
                  debugger;


                  for (count = 0; count < this.state.StateChartItemGUIDNode.length; count++) {

                    if (!this.userExists(this.state.StateChartItemGUIDNode[count]["EmpEmail"])) {
                      orgChartNodes.push(new ChartItem(

                        this.state.StateChartItemGUIDNode[count]["guid"],
                        this.state.StateChartItemGUIDNode[count]["title"],
                        this.state.StateChartItemGUIDNode[count]["Url"],
                        //orgChartItems[count]["Url"],
                        //orgChartItems[count].Managers ? orgChartItems[count].Managers.ID : undefined,
                        this.state.StateChartItemGUIDNode[count]["parent_guid"],
                        this.state.StateChartItemGUIDNode[count]["JobTitle"],
                        this.state.StateChartItemGUIDNode[count]["Manager"],
                        this.state.StateChartItemGUIDNode[count]["Department"],
                        this.state.StateChartItemGUIDNode[count]["EmpEmail"]

                      ));
                    }
                    // let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
                    // orgChartHoldNodesArrow.push(new HoldChartItemArrow(
                    // `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
                  }
                  debugger;


                  this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
                  //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
                  this.setState({ OrgChartNodesArray: orgChartNodes });
                  console.log("2: " + this.state.OrgChartNodesArray);
                  //this.CountDepthLevel(orgChartNodes);

                  var arrayToTree: any = require('array-to-tree');
                  debugger;

                  var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);
                  debugger;


                  var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
                  this.setState({
                    orgChartItems: JSON.parse(output)
                  });
                  debugger;
                });
            });

            //this.setState({ ManagerSetIdForExpand: managerid });



            let relevantResults: any = response.value;
            resolve(response.value);
          }, (error: any): void => {
            reject(error);
          });
      });
  }


  private async readOrgChartItemsArray(managerEmailID): Promise<IOrgChartItem[]> {


    let managerEmailID1 = "charan.sodhi@evolvous.com";
    return await new Promise<IOrgChartItem[]>(
      (resolve: (itemId: IOrgChartItem[]) => void, reject: (error: any) => void): void => {
        this.props.spHttpClient.get("" + this.props.pageContext.web.absoluteUrl
          + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"
          + encodeURIComponent('' + managerEmailID + '') + "'",
          // this.props.spHttpClient.get(
          //   `${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?$top 2&$select=Title,Id,EmpEmail,JobTitle,Department,Managers,Managers/ID&$expand=Managers/ID&$orderby=Title asc&$filter=Managers/ID eq ${managerid}`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
            // }).then((response: SPHttpClientResponse): Promise<{ value: IOrgChartItem[] }> => {
            //   
            //   return response.json();
            // })
            // .then((response: { value: IOrgChartItem[] }): void => {
            //   
          }).then((response) => {

            return response.json();
          })
          .then((response): void => {


            let imgURL: string;
            let userName: string;
            let guid: string;
            let parent_guid: number;
            let jobTitle: string;
            let emailID: string;
            let managers: string;
            let department: string;
            let ID: number;
            var properties = response.UserProfileProperties;
            let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
            let ChartItemGUIDNode: Array<ChartItemGUID> = [];
            var count: number;

            for (var i = 0; i < properties.length; i++) {

              var property = properties[i];
              if (property.Key == "WorkEmail") {
                emailID = property.Value;
              }
              if (property.Key == "WorkEmail") {
                imgURL = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${property.Value}`;
              }
              if (property.Key == "PreferredName") {
                userName = response.DisplayName;
              }
              if (property.Key == "UserProfile_GUID") {
                guid = property.Value;
              }
              if (property.Key == "SPS-JobTitle") {
                jobTitle = property.Value;
              }
              if (property.Key == "Department") {
                department = property.Value;
              }
              if (property.Key == "SPS-SharePointHomeExperienceState") {
                ID = property.Value;
              }
            }

            parent_guid = this.state.StateChartItemGUIDNode[0].guid;
            ChartItemGUIDNode = this.state.StateChartItemGUIDNode;

            ChartItemGUIDNode.push(new ChartItemGUID(
              guid,
              userName,
              imgURL,
              ID,
              parent_guid,
              jobTitle,
              this.state.StateChartItemGUIDNode[0].Manager,
              department,
              emailID)
            );


            // this.state.orgChartHoldNodesArrow.push(ChartItemGUIDNode);


            this.setState({ StateChartItemGUIDNode: ChartItemGUIDNode });
            //this.setState({ StateChartItemGUIDNode: orgChartHoldNodesArrow });
            //this.setState({ StateChartItemGUIDNode: ChartItemGUIDNode });

            //var properties = response.UserProfileProperties;
            var peers = response.Peers;


            //this.setState({ ManagerSetIdForExpand: managerEmailID });
            this.setState({ ManagerSetIdForExpandEmailID: this.state.StateChartItemGUIDNode[0].EmpEmail });


            let relevantResults: any = response.value;
            resolve(response.value);
          }, (error: any): void => {
            reject(error);
          });
      });
  }

  private processOrgChartItems(): void {

    this.readOrgChartItemsID().then((orgChartItemsOutter: IOrgChartItem[]) => {

      let peers = orgChartItemsOutter["DirectReports"];

      peers.forEach((elementRow) => {

        this.readOrgChartItems(elementRow).then((orgChartItemsInner: IOrgChartItem[]) => {



          // if (orgChartItemsOutter.length > 0) {
          //   orgChartNodes.push(new ChartItem(
          //     orgChartItemsOutter[0].Id,
          //     orgChartItemsOutter[0]["Title"],
          //     "https://evolvous.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?userId=accounts@evolvous.com",
          //     undefined,
          //     orgChartItemsOutter[0]["JobTitle"],
          //     orgChartItemsOutter[0]["Manager"],
          //     orgChartItemsOutter[0]["Department"],
          //     orgChartItemsOutter[0]["EmpEmail"]));

          //   if (orgChartItems.length > 1)
          //     orgChartHoldNodesArrow.push(new HoldChartItemArrow(
          //       `#CollapseContentSingle${orgChartItemsOutter[0].Id}`, "StockDown"));
          //   else {
          //     orgChartHoldNodesArrow.push(new HoldChartItemArrow(
          //       `#CollapseContentSingle${orgChartItemsOutter[0].Id}`, "StockDown"));
          //   }
          //   //this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
          // }

          // peers.forEach((elementEmail) => {

          //     this.readOrgChartItemsArray(elementEmail)
          //       .then((response) => {

          //         JSON.stringify(response);
          //         return JSON.stringify(response);
          //       }).then((responseData) => {

          //         var properties = responseData

          //       });
          //   });
          let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
          let orgChartNodes: Array<ChartItem> = [];
          var count: number;
          for (count = 0; count < this.state.StateChartItemGUIDNode.length; count++) {


            orgChartNodes.push(new ChartItem(
              this.state.StateChartItemGUIDNode[count]["id"],
              this.state.StateChartItemGUIDNode[count]["title"],
              this.state.StateChartItemGUIDNode[count]["Url"],
              this.state.StateChartItemGUIDNode[count]["parent_id"],
              this.state.StateChartItemGUIDNode[count]["JobTitle"],
              this.state.StateChartItemGUIDNode[count]["Manager"],
              this.state.StateChartItemGUIDNode[count]["Department"],
              this.state.StateChartItemGUIDNode[count]["EmpEmail"]
            ));


            //let orgChartHoldNodesArrow: Array<HoldChartItemArrow> = [];
            // orgChartHoldNodesArrow.push(new HoldChartItemArrow(
            //   `#CollapseContentSingle${orgChartItems[count].Id}`, "StockDown"));
          }
          debugger;
          this.setState({ StockDownUserIdValue: orgChartHoldNodesArrow });
          //this.setState({ LoadUpNextUser: orgChartItemsOutter[0]["EmpEmail"] });
          this.setState({ OrgChartNodesArray: orgChartNodes });
          console.log("2: " + this.state.OrgChartNodesArray);
          //this.CountDepthLevel(orgChartNodes);

          var arrayToTree: any = require('array-to-tree');
          debugger;
          var orgChartHierarchyNodes: any = arrayToTree(orgChartNodes);
          debugger;

          var output: any = JSON.stringify(orgChartHierarchyNodes[0]);
          this.setState({
            orgChartItems: JSON.parse(output)
          });
          debugger;
          //this.setState({ loadUpNext: false });
          //this.pow(orgChartHierarchyNodes[0]);
        });
      });
    });

  }

  private _renderLimitedSearch() {
    return (
      <CompactPeoplePicker
        onChange={this._onChange.bind(this)}
        onResolveSuggestions={this._onFilterChanged}
        getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
        pickerSuggestionsProps={suggestionProps}
        className={'ms-PeoplePicker'}
        key={'normal'}
      />
    );
  }

  private async _onChange(items: any[]) {
    await this.setState({
      selectedItems: items
    });

    await this.setState({ textSearchUserValue: (items.length > 0 ? items[0].EMail : "") });
    console.log("Textbox value:1 " + this.state.textSearchUserValue);
    if (this.props.onChange) {
      this.props.onChange(items);
    }

    this.processOrgChartItems();
  }
  // @autobind
  private _onFilterChanged = (filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) => {
    let _peopleList1 = [];
    if (filterText) {
      if (filterText.length > 0) {

        console.log("Textbox value:2 " + this.state.textSearchUserValue);
        if (this.state.textSearchUserValue.length > 0) {
          return [];// this.SearchPeople(filterText);//_peopleList1;
        }
        else {
          return this.SearchPeople(filterText);//_peopleList1;
        }
      }
    } else {
      return [];
    }
  }
  /**
   * @function
   * Returns people results after a REST API call
   */
  // public SearchPeople(terms: string): Promise<IPersonaProps[]> {

  //   return new Promise<IPersonaProps[]>((resolve, reject) =>
  //     //fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'`,
  //     fetch(`${this.props.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items?$select=ID,Title,JobTitle,Manager,Department,EmpEmail&$orderby=Title asc&$filter=startswith(Title,'${terms}')`,
  //       {
  //         headers: {
  //           'Accept': 'application/json;odata=nometadata',
  //           'odata-version': ''
  //         }
  //       })
  //       .then((response) => {
  //         return response.json();
  //       })
  //       .then((response: { value: IOrgChartItem[] }): void => {

  //         let relevantResults: any = response.value;
  //         //let resultCount: number = relevantResults.TotalRows;
  //         let people1 = [];
  //         let count = 0;

  //         relevantResults.forEach((row) => {

  //           let persona: IPersonaProps = {};

  //           persona.secondaryText = row["Title"];
  //           persona.imageUrl = `/_vti_bin/DelveApi.ashx/people/profileimage?userId=dev1@evolvous.com`;
  //           persona.EMail = row["EmpEmail"];
  //           persona.primaryText = row["Title"];

  //           people1.push(persona);
  //         });

  //         // this.setState({ personaState: people });

  //         resolve(people1);
  //       }, (error: any): void => {
  //         reject(this._peopleList = []);
  //       }));
  // }


  public SearchPeople(terms: string): Promise<IPersonaProps[]> {

    return new Promise<IPersonaProps[]>((resolve, reject) =>
      fetch(`/_api/search/query?querytext='(accountname:*${terms}*)'&rowlimit=10&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'&selectproperties=
    'LanguageProficiency,WorkPhone,PictureURL,PreferredName,Country,Skills,Department,
    Language,JobTitle,Path,WorkEmail,accountName,Manager,ID'&sortlist='FirstName:ascending'`,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response) => {
          return response.json();
        })
        .then((response: { PrimaryQueryResult: IPeopleDataResult }): void => {
          let relevantResults: any = response.PrimaryQueryResult.RelevantResults;
          let resultCount: number = relevantResults.TotalRows;
          let people1 = [];
          if (resultCount > 0) {
            relevantResults.Table.Rows.forEach((row) => {
              let persona: IPersonaProps = {};
              row.Cells.forEach((cell) => {
                //person[cell.Key] = cell.Value; 
                if (cell.Key === 'JobTitle')
                  persona.secondaryText = cell.Value;
                if (cell.Key === 'WorkEmail')
                  persona.imageUrl = `${profileURL}/_vti_bin/DelveApi.ashx/people/profileimage?userId=${cell.Value}`;
                if (cell.Key === 'PreferredName')
                  persona.primaryText = cell.Value;
                if (cell.Key === 'WorkEmail') {
                  persona.EMail = cell.Value == "" ? "Null" : cell.Value;
                }
              });
              people1.push(persona);
            });


            // this.setState({ personaState: people });
          }
          resolve(people1);
        }, (error: any): void => {
          reject(this._peopleList = []);
        }));
  }
}
