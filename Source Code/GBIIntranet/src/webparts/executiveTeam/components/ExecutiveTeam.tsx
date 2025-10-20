import * as React from 'react';
import styles from './ExecutiveTeam.module.scss';
import { IExecutiveTeamProps } from './IExecutiveTeamProps';
import { escape } from '@microsoft/sp-lodash-subset';

require('../assets/style.css');

export default class ExecutiveTeam extends React.Component<IExecutiveTeamProps, {}> {
  public render(): React.ReactElement<IExecutiveTeamProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
     <section className="team">
    <h2>Executive Team</h2>
    <div className="team-grid">
      <div className="team-card">
        <img src={require('../assets/user.png')} alt="Team Member"/>
        <div className="overlay">
          <h4>Darlene Robertson</h4>
          <p>Marketing Manager</p>
        </div>
      </div>

      <div className="team-card">
        <img src={require('../assets/user2.png')} alt="Team Member"/>
        <div className="overlay">
          <h4>Darlene Robertson</h4>
          <p>Marketing Manager</p>
        </div>
      </div>

      <div className="team-card">
        <img src={require('../assets/user.png')} alt="Team Member"/>
        <div className="overlay">
          <h4>Darlene Robertson</h4>
          <p>Marketing Manager</p>
        </div>
      </div>

      <div className="team-card">
        <img src={require('../assets/user2.png')} alt="Team Member"/>
        <div className="overlay">
          <h4>Darlene Robertson</h4>
          <p>Marketing Manager</p>
        </div>
      </div>
    </div>
  </section>

    );
  }
}
