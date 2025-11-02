import * as React from "react";
import styles from "./EmployeeDirectory.module.scss";
import { IEmployeeDirectoryProps } from "./IEmployeeDirectoryProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { PersonaCard } from "../components/PersonaCard";
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize, PrimaryButton, Spinner, SpinnerSize, TextField } from "@fluentui/react";
import { Pagination } from "./Pagination";
require("../assets/css/fabric.min.css");
require("../assets/css/style.css");
import { SPHttpClient, SPHttpClientResponse, MSGraphClient, AadHttpClient } from "@microsoft/sp-http";
import { Profiles, sp, Web } from "@pnp/sp/presets/all";
import * as $ from "jquery";

const az: string[] = ["See All", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

export interface IEmployeeDirectoryState {
  allUsers: any;
  profileProperties: any;
  isLoading: boolean;
  searchArray: any;
  indexSelectedKey: any;
  searchText: any;
  currentPage: number;
  totalPages: number;
  items: any;
  allEmpData: any;
}

const viewCount = 12;

export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeDirectoryState> {
  private graphClient: MSGraphClient = null;

  constructor(props: IEmployeeDirectoryProps, state: IEmployeeDirectoryState) {
    super(props);
    this.state = {
      allUsers: [],
      profileProperties: [],
      isLoading: true,
      searchArray: [],
      indexSelectedKey: "See All",
      searchText: "",
      currentPage: 1,
      totalPages: 5,
      items: [],
      allEmpData: [],
    };
    this._selectedIndex = this._selectedIndex.bind(this);
  }
  public render(): React.ReactElement<IEmployeeDirectoryProps> {
    return (
      <>
        <div className="ms-Grid directory-container">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12" style={{ textAlign: "center" }}>
              <h2 className="EmpCustomHeader">
                <div>Employee Directory</div>
                <PrimaryButton
                  className="refreshBTN"
                  onClick={() => {
                    this.setState({ isLoading: true }, () => {
                      this.getUserDirectory();
                    });
                  }}
                >
                  Refresh Data
                </PrimaryButton>
              </h2>
            </div>
            <div className="ms-Grid-col ms-sm12" style={{ textAlign: "center" }}></div>
            <div className="ms-Grid-col ms-sm12">
              <TextField
                placeholder="Search"
                value={this.state.searchText}
                onChange={(event) => {
                  this.searchUsers(event.target["value"]);
                }}
              />
            </div>
            <div className="ms-Grid-col ms-sm12">
              <Pivot
                styles={{
                  root: {
                    paddingLeft: 10,
                    paddingRight: 10,
                    whiteSpace: "normally",
                  },
                }}
                linkFormat={PivotLinkFormat.tabs}
                selectedKey={this.state.indexSelectedKey}
                onLinkClick={this._selectedIndex}
                linkSize={PivotLinkSize.normal}
              >
                {az.map((index: string) => {
                  return <PivotItem headerText={index} itemKey={index} key={index} />;
                })}
              </Pivot>
            </div>

            {this.state.isLoading && (
              <div className="ms-Grid-col ms-sm12">
                <Spinner size={SpinnerSize.large} label={"loading ..."} />
              </div>
            )}
            <div className="ms-Grid-col ms-sm12 listingRow">
              {this.state.allEmpData.length > 0
                ? this.state.allEmpData.map((profile) => {
                    return <PersonaCard context={this.props.spfxContext} profileProperties={profile}></PersonaCard>;
                  })
                : this.state.isLoading == false && (
                    <div className="not-found" style={{ textAlign: "center" }}>
                      <p>No User Found</p>
                    </div>
                  )}
            </div>
            {this.state.isLoading == false && this.state.allEmpData.length > 0 ? (
              <div className="ms-Grid-col ms-sm12 list-paging">
                <Pagination
                  currentPage={this.state.currentPage}
                  totalPages={this.state.totalPages}
                  onChange={(page) => this.pagination(page, this.state.items)}
                  limiter={3} // Optional - default value 3
                />
              </div>
            ) : (
              <div></div>
            )}
          </div>
        </div>
      </>
    );
  }

  // function to get users from Azure AD
  public async componentDidMount(): Promise<void> {
    // this.graphClient = await this.props.spfxContext.msGraphClientFactory.getClient();
    // await this.graphClient.api("users").version("v1.0").select("id,mail,userPrincipalName,displayName").top(999).get().then((res) => {
    //   let next = res["@odata.nextLink"];
    //   this.setState({ allUsers: res.value });
    //   next ? this.getUserFromNextCalls(next) : this.getDataFromUserProfile();
    // }).catch((err) => {
    //   console.log(err);
    // });

    await this.CheckLoginUserPermission();
    await this.getJSONFile();
  }

  public getUserDirectory = async () => {
    this.graphClient = await this.props.spfxContext.msGraphClientFactory.getClient();
    await this.graphClient
      .api("users")
      .version("v1.0")
      .select("id,mail,userPrincipalName,displayName")
      .top(999)
      .get()
      .then((res) => {
        let next = res["@odata.nextLink"];
        this.setState({ allUsers: res.value });
        next ? this.getUserFromNextCalls(next) : this.getDataFromUserProfile();
      })
      .catch((err) => {
        console.log(err);
      });
  };

  public CheckLoginUserPermission() {
    $.ajax({
      url: this.props.siteUrl + "/_api/web/currentUser/issiteadmin",
      method: "GET",
      async: true,
      headers: {
        Accept: "application/json;odata=verbose",
      },
      success: function (data) {
        data.d.IsSiteAdmin == true ? $(".refreshBTN").show() : $(".refreshBTN").hide();
      },
      error: function (data) {
        console.log("Could not find Title");
      },
    });
  }

  // function to get more then 1000 users from Azure AD
  public getUserFromNextCalls(url) {
    this.props.spfxContext.aadHttpClientFactory
      .getClient("https://graph.microsoft.com")
      .then((client: AadHttpClient) => {
        return client.get(url, AadHttpClient.configurations.v1);
      })
      .then((response) => {
        return response.json();
      })
      .then((json) => {
        let temp = this.state.allUsers;
        json.value.forEach((element) => {
          temp.push(element);
        });
        this.setState({ allUsers: temp });
        var next = json["@odata.nextLink"];
        next ? this.getUserFromNextCalls(next) : this.getDataFromUserProfile();
      });
  }

  // function to get users properties from user profile
  public getDataFromUserProfile = async () => {
    let profileProperties = [],
      count = 0;
    if (this.state.allUsers.length > 0) {
      for (const ele of this.state.allUsers) {
        try {
          let mail = ele.userPrincipalName.toLowerCase();
          let loginName = "i:0#.f|membership|" + mail;

          // let loginName = "i:0#.f|membership|" + ele.userPrincipalName;
          await sp.profiles
            .getPropertiesFor(loginName)
            .then((data) => {
              let dep = "",
                delve = "",
                office = "",
                image = "",
                cellphone = "";
              image = this.props.siteUrl + "/_layouts/15/userphoto.aspx?size=L&username=" + data.Email;

              if (data && data.UserProfileProperties && Array.isArray(data.UserProfileProperties)) {
                data.UserProfileProperties.forEach((x) => {
                  if (x.Key === "Department") dep = x.Value;
                  if (x.Key === "msOnline-ObjectId") delve = x.Value;
                  if (x.Key === "Office") office = x.Value;
                  if (x.Key === "CellPhone") cellphone = x.Value;
                });
              }

              profileProperties.push({
                Name: data.DisplayName,
                email: data.Email,
                job_title: data.Title,
                Department: dep,
                DelveId: delve,
                Office_location: office,
                Image: image,
                m_phone: cellphone,
              });
              count = count + 1;
            })
            .catch((err) => {
              console.log(err);
            });
          if (this.state.allUsers.length == count) {
            this.storeJsonData(JSON.stringify(profileProperties));
            // let sortData = profileProperties.sort(this.SortDate);
            // this.setState({ profileProperties: sortData, searchArray: sortData, isLoading: false });
            // this.pagination(this.state.currentPage, sortData);
          }
        } catch (err) {
          console.log(err);
        }
      }
      // this.state.allUsers.forEach(async (ele, i) => {

      //   // setTimeout(async () => {
      //   let loginName = "i:0#.f|membership|" + ele.userPrincipalName;
      //   await sp.profiles.getPropertiesFor(loginName).then((data) => {
      //     let dep = "", delve = "", office = "", image = "", cellphone = "";
      //     image = this.props.siteUrl + "/_layouts/15/userphoto.aspx?size=L&username=" + data.Email;
      //     data.UserProfileProperties.filter(x => { if (x.Key == "Department") { dep = x.Value; } });
      //     data.UserProfileProperties.filter(x => { if (x.Key == "msOnline-ObjectId") { delve = x.Value; } });
      //     data.UserProfileProperties.filter(x => { if (x.Key == "Office") { office = x.Value; } });
      //     data.UserProfileProperties.filter(x => { if (x.Key == "CellPhone") { cellphone = x.Value; } });

      //     profileProperties.push({
      //       "DisplayName": data.DisplayName,
      //       "Email": data.Email,
      //       "JobTitle": data.Title,
      //       "Department": dep,
      //       "DelveUrl": delve,
      //       "OfficeLocation": office,
      //       "ImgUrl": image,
      //       "CellPhone": cellphone,
      //     });
      //     count = count + 1;
      //   }).catch((err) => {
      //     console.log(err);
      //   });
      //   if (this.state.allUsers.length == count) {
      //     let sortData = profileProperties.sort(this.SortDate);
      //     this.setState({ profileProperties: sortData, searchArray: sortData, isLoading: false });
      //     this.pagination(this.state.currentPage, sortData);
      //   }
      // });
    }
  };

  public getJSONFile = async () => {
    var cnt1 = 0;
    var webAbsoluteUrl = this.props.spfxContext.pageContext.site.absoluteUrl;
    var siteURL = webAbsoluteUrl.split(window.location.hostname)[1];
    var docLibraryName = "SiteAssets/UserProfileData/profile.json";
    var destinationURL = siteURL + "/" + docLibraryName;
    let web = Web(webAbsoluteUrl);
    var x = await web
      .getFileByServerRelativeUrl(destinationURL)
      .getJSON()
      .then((res) => {
        let UserArr = [];
        if (res.length > 0) {
          res.map((ele) => {
            let userJson = {};
            userJson["DisplayName"] = ele.Name ? ele.Name : "";
            userJson["JobTitle"] = ele.job_title ? ele.job_title : "";
            userJson["OfficeLocation"] = ele.Office_location ? ele.Office_location : "";
            userJson["Department"] = ele.Department ? ele.Department : "";
            userJson["Phone"] = ele.b_phone ? ele.b_phone : "";
            userJson["CellPhone"] = ele.m_phone ? ele.m_phone : "";
            userJson["Email"] = ele.email ? ele.email : "";
            userJson["ImgUrl"] = ele.Image ? ele.Image : "";
            userJson["DelveUrl"] = ele.DelveId ? ele.DelveId : "";

            UserArr.push(userJson);
          });
          let sortData = UserArr.sort(this.SortDate);
          this.setState({ profileProperties: sortData, searchArray: sortData, isLoading: false });
          this.pagination(this.state.currentPage, sortData);
          // this.setState({ userData: sortData, searchArray: sortData });
        }
      })
      .catch((er) => {
        console.log(er);
      });
  };

  public storeJsonData = async (json) => {
    json = JSON.parse(json);
    var blob = new Blob([JSON.stringify(json)], {
      type: "text/plain",
    });
    var fileName = "profile.json";
    var webAbsoluteUrl = this.props.spfxContext.pageContext.site.absoluteUrl;
    var siteURL = webAbsoluteUrl.split(window.location.hostname)[1];
    var docLibraryName = "SiteAssets/UserProfileData/";
    var destinationURL = siteURL + "/" + docLibraryName;

    var arrayBuffer;
    var fileReader = new FileReader();
    var fileInfos = [];
    var Mythis = this;
    fileReader.onload = async function () {
      arrayBuffer = this.result;
      let web = Web(webAbsoluteUrl);
      var x = await web
        .getFolderByServerRelativeUrl(destinationURL)
        .files.add(fileName, arrayBuffer, true)
        .then((res) => {
          Mythis.setState({ isLoading: false }, () => {
            Mythis.getJSONFile();
          });
        })
        .catch((e) => {
          console.log(e);
          Mythis.setState({ isLoading: false });
        });
    };
    fileReader.readAsArrayBuffer(blob);
  };

  // function for sort date
  private SortDate = (a, b) => {
    if (a.DisplayName != "" && b.DisplayName != "") {
      let dateA = a.DisplayName.toLowerCase();
      let dateB = b.DisplayName.toLowerCase();
      return dateA > dateB ? 1 : -1;
    }
  };

  // function to set pagination
  public pagination(crntPage, libraryData) {
    var startCount = (crntPage - 1) * viewCount;
    var endCount = crntPage * viewCount;
    let pagedArr = libraryData.slice(startCount, endCount);
    this.setState({
      currentPage: 1,
      items: libraryData,
      totalPages: Math.ceil(libraryData.length / viewCount),
    });
    this.mapPageData(pagedArr);
  }

  // function to set or map data to pages
  public async mapPageData(pageData: any[]) {
    this.setState({ allEmpData: pageData });
  }

  // function to search users based on alphabet selections
  private _selectedIndex(item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) {
    this.setState({ indexSelectedKey: item.props.itemKey, searchText: "" }, () => {
      if (item.props.itemKey != "See All") {
        let Search = this.state.searchArray.filter((value) => {
          let val = value.DisplayName.toLowerCase();
          if (val.charAt(0) == item.props.itemKey.toLowerCase()) {
            return value;
          }
        });
        this.setState({ profileProperties: Search });
        this.pagination(this.state.currentPage, Search);
      } else {
        this.setState({ profileProperties: this.state.searchArray });
        this.pagination(this.state.currentPage, this.state.searchArray);
      }
    });
  }

  // function to search users
  public searchUsers = (text) => {
    this.setState({ searchText: text, indexSelectedKey: "See All" });
    let SearchUser = this.state.searchArray.filter((value) => {
      let val = value.DisplayName.toLowerCase();
      let val1 = value.Department.toLowerCase();
      if (val.includes(text.toLowerCase()) || val1.includes(text.toLowerCase())) {
        return value;
      }
    });
    this.setState({ profileProperties: SearchUser });
    this.pagination(this.state.currentPage, SearchUser);
  };
}
