import * as React from 'react';
import styles from './SocialMedia.module.scss';
import { ISocialMediaProps } from './ISocialMediaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import InstagramEmbed from "./Insta";
import FacebookPageEmbed from "./Facebook";

require('../assets/fabric.min.css');
require("../assets/style.css");

export default class SocialMedia extends React.Component<ISocialMediaProps, {}> {
  public render(): React.ReactElement<ISocialMediaProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section>
        <div className="Stay_Connected">
          <h2 className="Stay_Connected_title">Stay Connected</h2>
          <div className="p-15">
            {/* <iframe style={{ border: "none" }} height={"420px"} width={"100%"} data-tweet-url="https://www.linkedin.com/company/netuno-usa-inc" src="data:text/html;charset=utf-8,%3Ca%20class%3D%22twitter-timeline%22%20href%3D%22https%3A//twitter.com/iZOOlogic%3Fref_src%3Dtwsrc%255Etfw%22%3ETweets%20by%20iZOOlogic%3C/a%3E%0A%3Cscript%20async%20src%3D%22https%3A//platform.twitter.com/widgets.js%22%20charset%3D%22utf-8%22%3E%3C/script%3E%0A"></iframe> */}
            <div className="ms-Grid">
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <InstagramEmbed />
                </div>
                 <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                     <div className="ms-Grid-row">
                 <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                   <video controls>
              <source src="https://bullioninternational.sharepoint.com/sites/intranet/Videos/video1.mp4" type="video/mp4" />
            </video>
                 </div>
                 <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                   <video controls>
              <source src="https://bullioninternational.sharepoint.com/sites/intranet/Videos/video1.mp4" type="video/mp4" />
            </video>
                 </div>
                              <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                   <video controls>
              <source src="https://bullioninternational.sharepoint.com/sites/intranet/Videos/video1.mp4" type="video/mp4" />
            </video>
                 </div>

                     </div>
                 </div>
                {/* <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <FacebookPageEmbed />
                </div> */}
              </div>
            </div>
          </div>
        </div>
      </section>
    );
  }
}
