import * as React from "react";
import styles from "./ExecutiveTeam.module.scss";
import { IExecutiveTeamProps } from "./IExecutiveTeamProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";

require("../assets/style.css");
export interface IExecutiveTeamState {
  executiveTeams: any;
}
export default class ExecutiveTeam extends React.Component<IExecutiveTeamProps, IExecutiveTeamState> {
  constructor(props: IExecutiveTeamProps, state: IExecutiveTeamState) {
    super(props);
    this.state = {
      executiveTeams: [],
    };
  }
  public render(): React.ReactElement<IExecutiveTeamProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <section id="ExecutiveTeam" className="team">
        <h2>Executive Team</h2>
        <div className="team-grid">
          {this.state.executiveTeams.length > 0 &&
            this.state.executiveTeams.map((ele, ind) => {
              let imageURL = ele.AttachmentFiles.length > 0 ? ele.AttachmentFiles[0].ServerRelativeUrl : ele.Photo ? JSON.parse(ele.Photo).serverRelativeUrl : require(`../assets/user.png`);
              return (
                <div className="team-card">
                  <img src={imageURL} alt="Team Member" />
                  <div className="overlay">
                    <h4>{ele.Title}</h4>
                    <p>{ele.Designation}</p>
                  </div>
                </div>
              );
            })}
          {/* <div className="team-card">
            <img src={require("../assets/user.png")} alt="Team Member" />
            <div className="overlay">
              <h4>Darlene Robertson</h4>
              <p>Marketing Manager</p>
            </div>
          </div>

          <div className="team-card">
            <img src={require("../assets/user2.png")} alt="Team Member" />
            <div className="overlay">
              <h4>Darlene Robertson</h4>
              <p>Marketing Manager</p>
            </div>
          </div>

          <div className="team-card">
            <img src={require("../assets/user.png")} alt="Team Member" />
            <div className="overlay">
              <h4>Darlene Robertson</h4>
              <p>Marketing Manager</p>
            </div>
          </div>

          <div className="team-card">
            <img src={require("../assets/user2.png")} alt="Team Member" />
            <div className="overlay">
              <h4>Darlene Robertson</h4>
              <p>Marketing Manager</p>
            </div>
          </div> */}
        </div>
      </section>
    );
  }
  public componentDidMount = async () => {
    await this.getExecutiveTeamsDetails();
  }

  // get executive teams details from Executive Teams sharepoint list
  private getExecutiveTeamsDetails = async () => {
    try {
      const ExecutiveTeamsDetails = await sp.web.lists.getByTitle("Executive Team").items.select("Title,ID,Designation,Photo").expand("AttachmentFiles").get();

      if (ExecutiveTeamsDetails.length > 0) {
        this.setState({ executiveTeams: ExecutiveTeamsDetails });
      }
    } catch (error) {
      console.log(error);
    }
  }
}
