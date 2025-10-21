import * as React from "react";
import styles from "./VisionMission.module.scss";
import { IVisionMissionProps } from "./IVisionMissionProps";
import { escape } from "@microsoft/sp-lodash-subset";

require("../assets/style.css");

export default class VisionMission extends React.Component<IVisionMissionProps, {}> {
  public render(): React.ReactElement<IVisionMissionProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <section className="vision-mission">
        <div className="card">
          <div className="icon">
            <img src={require("../assets/vision.png")} />
          </div>
          <h3>{this.props.visiontitle}</h3>
          <p>{this.props.visiondescription}</p>
        </div>
        <div className="card">
          <div className="icon">
            <img src={require("../assets/mission.png")} />
          </div>
          <h3>{this.props.missiontitle}</h3>
          <p>{this.props.missiondescription}</p>
        </div>
      </section>
    );
  }
}
