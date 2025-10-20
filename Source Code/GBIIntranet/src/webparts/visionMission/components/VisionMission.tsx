import * as React from 'react';
import styles from './VisionMission.module.scss';
import { IVisionMissionProps } from './IVisionMissionProps';
import { escape } from '@microsoft/sp-lodash-subset';

require('../assets/style.css');

export default class VisionMission extends React.Component<IVisionMissionProps, {}> {
  public render(): React.ReactElement<IVisionMissionProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
     <section className="vision-mission">
    <div className="card">
      <div className="icon"><img src={require('../assets/vision.png')} /></div>
      <h3>Vision</h3>
      <p>To be the most trusted brand in physical precious metals investment, with an unwavering focus on superior products, comprehensive education, and exceptional service.</p>
    </div>
    <div className="card">
      <div className="icon"><img src={require('../assets/mission.png')} /></div>
      <h3>Mission</h3>
      <p>Empowering the world to preserve and grow their wealth through informed investing and seamless access to physical precious metals and other tangible assets.</p>
    </div>
  </section>
    );
  }
}
