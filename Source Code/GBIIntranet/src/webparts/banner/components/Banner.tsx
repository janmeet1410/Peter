import * as React from "react";
import styles from "./Banner.module.scss";
import { IBannerProps } from "./IBannerProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import * as moment from "moment";

require("../assets/style.css");
export interface IBannerState {
  quickLinks: any;
}
export default class Banner extends React.Component<IBannerProps, IBannerState> {
  constructor(props: IBannerProps, state: IBannerState) {
    super(props);
    this.state = {
      quickLinks: [],
    };
  }
  public render(): React.ReactElement<IBannerProps> {
    return (
      <>
        <section className="hero-section">
          <div className="hero-content">
            <p className="date">{moment().format("D MMM YYYY").toUpperCase()}</p>
            {/* <h1>
        Discover GBI's Leading<br />
        Institutional Grade Physical<br />
        Precious Metals Platform.
      </h1> */}
            <h1>{this.props.bannerdescription}</h1>
          </div>
          <div className="hero-image">
            <img src={require("../assets/Images/sideimg.png")} alt="Business Woman" />
          </div>
        </section>
        <section className="card-section quicklinks">
          {this.state.quickLinks.length > 0 &&
            this.state.quickLinks.map((ele, ind) => {
              let imageURL = ele.AttachmentFiles.length > 0 ? ele.AttachmentFiles[0].ServerRelativeUrl : ele.Icon ? JSON.parse(ele.Icon).serverRelativeUrl : require(`../assets/Images/computer.png`);
              return (
                <a target="_blank" data-interception="off" href={ele.Link ? ele.Link.Url : "#"} className="card">
                  <img src={imageURL} alt={ele.Title} />
                  <p>{ele.Title}</p>
                </a>
              );
            })}
          {/* <div className="card">
            <img src={require("../assets/Images/computer.png")} alt="IT help" />
            <p>IT help</p>
          </div>
          <div className="card">
            <img src={require("../assets/Images/financing.png")} alt="Finance" />
            <p>Finance</p>
          </div>
          <div className="card">
            <img src={require("../assets/Images/hr.png")} alt="HR" />
            <p>HR</p>
          </div>
          <div className="card">
            <img src={require("../assets/Images/mortarboard.png")} alt="ADP" />
            <p>ADP</p>
          </div>
          <div className="card">
            <img src={require("../assets/Images/computer.png")} alt="Social wall" />
            <p>Manager's Corner</p>
          </div> */}
        </section>
      </>
    );
  }

  public componentDidMount = async () => {
    await this.getQuickLinksDetails();
  }

  // get quick links details from Quick Links sharepoint list
  private getQuickLinksDetails = async () => {
    try {
      const QuickLinkssDetails = await sp.web.lists.getByTitle("Quick Links").items.select("Title,Link,ID,Icon,LinkOrderBy").expand("AttachmentFiles").orderBy("LinkOrderBy", true).get();

      if (QuickLinkssDetails.length > 0) {
        this.setState({ quickLinks: QuickLinkssDetails });
      }
    } catch (error) {
      console.log(error);
    }
  }
}
