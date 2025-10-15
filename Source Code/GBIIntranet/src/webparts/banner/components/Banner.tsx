import * as React from 'react';
import styles from './Banner.module.scss';
import { IBannerProps } from './IBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

require('../assets/style.css');

export default class Banner extends React.Component<IBannerProps, {}> {
  public render(): React.ReactElement<IBannerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <>
       <section className="hero-section">
    <div className="hero-content">
      <p className="date">8 OCT 2025</p>
      <h1>
        Discover GBI's Leading<br />
        Institutional Grade Physical<br />
        Precious Metals Platform.
      </h1>
    </div>
    <div className="hero-image">
      <img src={require('../assets/Images/sideimg.png')} alt="Business Woman" />
    </div>
  </section>
       <section className="card-section quicklinks">
    <div className="card">
      <img src={require('../assets/Images/computer.png')} alt="IT help"/>
      <p>IT help</p>
    </div>
    <div className="card">
      <img src={require('../assets/Images/financing.png')} alt="Finance"/>
      <p>Finance</p>
    </div>
    <div className="card">
      <img src={require('../assets/Images/hr.png')} alt="HR"/>
      <p>HR</p>
    </div>
    <div className="card">
      <img src={require('../assets/Images/mortarboard.png')} alt="ADP"/>
      <p>ADP</p>
    </div>
    <div className="card">
      <img src={require('../assets/Images/computer.png')} alt="Social wall"/>
      <p>Manager's Corner</p>
    </div>
  </section>
      </>

    );
  }
}
