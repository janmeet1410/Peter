import * as React from 'react';
import styles from './Subbanner.module.scss';
import { ISubbannerProps } from './ISubbannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

require("../assets/style.css");

export default class Subbanner extends React.Component<ISubbannerProps, {}> {
  public render(): React.ReactElement<ISubbannerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="hero-section">
               <div className="hero-content">
                 {/* <p className="date">{moment().format("D MMM YYYY").toUpperCase()}</p> */}
                 {/* <h1>
             Discover GBI's Leading<br />
             Institutional Grade Physical<br />
             Precious Metals Platform.
           </h1> */}
                 <h1>{this.props.bannerdescription}</h1>
                 <p></p>
               </div>
               <div className="hero-image">
                 <img src={require("../assets/Images/sideimg.png")} alt="Business Woman" />
               </div>
             </section>
    );
  }
}
