import * as React from "react";
import styles from "./OrgChart.module.scss";
import { IOrgChartProps } from "./IOrgChartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import { ChartItem } from "./IUserItem";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";
import { SPComponentLoader } from "@microsoft/sp-loader";
import * as $ from "jquery";

// require('Jquery');
require("../assets/css/style.css");
/* Importing Js */
require("../assets/js/jquery.min.js");
// require('./js/jquery.orgchart.js');
require("../assets/js/jquery.orgchart.js");

SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.2/jquery.min.js");
SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/orgchart/3.4.0/css/jquery.orgchart.min.css");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/orgchart/3.5.0/js/jquery.orgchart.min.js");

interface OrgItem {
  Id: number;
  Title: string;
  JobTitle: string;
  PhotoUrl: string;
  ManagerId: number | null;
  children?: OrgItem[];
}
export interface IOrgChartState {
  orgChartItems: any;
  AllUsers: any;
  isLoading: any;
}
export default class OrgChart extends React.Component<IOrgChartProps, IOrgChartState> {
  constructor(props: IOrgChartProps, state: IOrgChartState) {
    super(props);
    this.state = {
      orgChartItems: [],
      AllUsers: [],
      isLoading: true,
    };
  }

  public render(): React.ReactElement<IOrgChartProps> {
    return (
      <section>
        <div id="org-chart" className="organizationChartViewer">
          <div className="container_43ab192f">
            {this.props.description.indexOf("@") != -1 ? (
              <>
                {this.state.isLoading ? (
                  <>
                    <Spinner size={SpinnerSize.large} label={"Please Wait ..."} />
                  </>
                ) : (
                  <>
                    <div className="canvas-div">
                      <div id="chart-container">
                        <div className="organization-chart-page-heading">
                          <h2 className="organization-chart-page-title">ORGANIZATION DIRECTORY</h2>
                          {/* <button className="expand-collapse-button" id="expand-btn">EXPAND ALL</button> */}
                        </div>
                      </div>
                    </div>
                  </>
                )}
              </>
            ) : (
              <>
                <h3>Please Configure webpart</h3>
              </>
            )}
          </div>
        </div>
      </section>
    );
  }

  public componentDidMount() {
    var reactHandler = this;
    reactHandler.GraphAPIUsers();
    if (!window.location.hash) {
      window.location.href = window.location + "#loaded";
      // window.location.reload();
    }
  }

  public getUserFromNextCall(url) {
    this.props.context.aadHttpClientFactory
      .getClient("https://graph.microsoft.com")
      .then((client: AadHttpClient) => {
        return client.get(url, AadHttpClient.configurations.v1);
      })
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        var temp = this.state.AllUsers;
        json.value.forEach((element) => {
          temp.push(element);
        });
        this.setState({ AllUsers: temp });
        var next = json["@odata.nextLink"];
        next ? this.getUserFromNextCall(next) : this.PassUsersToOrgChart();
      });
  }

  /*this function will check the parent is avaialble or not */
  public checkParentIsAvailable(arr, id) {
    var flag = false;
    var json = {};
    arr.forEach((element) => {
      if (element.id == id) {
        flag = true;
        json = element;
      }
    });

    var loop = true;
    var parent = json["parent_id"];
    /* if parent is available then we will check the hierarcy, if hierarchy is true then this user will pass to org chart. */
    while (loop && flag && parent != undefined) {
      // debugger;
      var data = parent ? this.verifyHierarchy(arr, parent) : "";

      if (data && data["parent_id"]) {
        parent = data["parent_id"];
      } else {
        if (data && data["parent_id"] == undefined) {
          loop = false;
        } else {
          flag = false;
        }
      }
      console.log(data);
    }
    return flag;
  }

  /*this function is used to check hierarchy */
  public verifyHierarchy(arr, id) {
    var json = false;
    arr.forEach((element) => {
      if (element.id == id) {
        json = element;
      }
    });
    return json;
  }

  public PassUsersToOrgChart() {
    let orgChartNodes: Array<ChartItem> = [];
    var output: any;

    // Map the JSON response to the output array

    if (this.state.AllUsers) {
      this.state.AllUsers.map((item: any) => {
        if (item.userPrincipalName && this.props.description.toLocaleLowerCase() == item.userPrincipalName.toLocaleLowerCase()) {
          orgChartNodes.push(new ChartItem(item.id, item.displayName, item.jobTitle, item.userPrincipalName, undefined, "/_layouts/15/userphoto.aspx?size=L&username=" + item.userPrincipalName));
        } else if (item.manager && item.manager.id && item.id != item.manager.id) {
          orgChartNodes.push(new ChartItem(item.id, item.displayName, item.jobTitle, item.userPrincipalName, item.manager ? item.manager.id : undefined, "/_layouts/15/userphoto.aspx?size=L&username=" + item.userPrincipalName));
        }
      });

      /*we have to remove those people whose manager do not have manager */
      let orgNodes: Array<ChartItem> = [];
      // this.validateHierarchy(orgChartNodes);
      orgChartNodes.forEach((element) => {
        var isManagerAvaileble = element.parent_id ? this.checkParentIsAvailable(orgChartNodes, element.parent_id) : "";
        if (isManagerAvaileble) {
          orgNodes.push(element);
        } else if (element.parent_id == undefined) {
          orgNodes.push(element);
        }
      });
      console.log("User Availeble", orgChartNodes);
      console.log("User Filtered", orgNodes);

      var arrayToTree: any = require("array-to-tree");
      var orgChartHierarchyNodes: any = arrayToTree(orgNodes);
      output = JSON.stringify(orgChartHierarchyNodes[0]); // orgChartHierarchyNodes[51]
      console.log(output);
      this.setState(
        {
          orgChartItems: JSON.parse(output),
          isLoading: false,
        },
        () => {
          this.bindOrgChart();
        }
      );
    }
  }

  private GraphAPIUsers = (): void => {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient) => {
      //client.api("users").select("displayName,mail,userPrincipalName,manager,givenName,jobTitle,mobilePhone,officeLocation,preferredLanguage,businessPhones,city,manager") // to get all users
      client
        .api("users")
        .expand("manager")
        .select("id,displayName,mail,userPrincipalName,manager,givenName,jobTitle,mobilePhone,officeLocation,preferredLanguage,businessPhones,city")
        .top(999)
        .get((err, res) => {
          if (err) {
            console.error(err);
            return;
          }
          if (this.props.description.indexOf("@") != -1) {
            var next = res["@odata.nextLink"];
            this.setState({ AllUsers: res.value });
            next ? this.getUserFromNextCall(next) : this.PassUsersToOrgChart();
          }
        });
    });
  };

  /* For Node Template */
  public nodeTemplate = (data) => {
    return `
       <div class="title"></div>
       <div class="content">
       <div class="profilePic"><img src="${data.url ? data.url : ""}" /></div>
       <div>
       <div class="empName" title='${data.name ? data.name : ""}' >${data.name ? data.name : ""}</div>
       <div class="empDesignation" title='${data.title ? data.title : ""}'>${data.title ? data.title : ""}</div>
       </div>
       </div>
    `;
    //  <div class="externalLink"><img src="${require("./images/external-link.png")}" /></div>
  };

  /* Function For Binding Org Chart */
  private bindOrgChart = (): void => {
    let myThis = this;
    $(document).ready(function () {
      var oc = $("#chart-container").orgchart({
        exportFilename: "Default  ",
        data: myThis.state.orgChartItems,
        nodeContent: "title",
        visibleLevel: 2,
        nodeTemplate: myThis.nodeTemplate,
      });
    });
  };
}
