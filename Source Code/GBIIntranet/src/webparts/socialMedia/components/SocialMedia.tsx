import * as React from "react";
import styles from "./SocialMedia.module.scss";
import { ISocialMediaProps } from "./ISocialMediaProps";
import { escape } from "@microsoft/sp-lodash-subset";
import InstagramEmbed from "./Insta";
import FacebookPageEmbed from "./Facebook";
import { sp } from "@pnp/sp/presets/all";

require("../assets/fabric.min.css");
require("../assets/style.css");
export interface ISocialMediaState {
  allVideos: any;
}
export default class SocialMedia extends React.Component<ISocialMediaProps, ISocialMediaState> {
  constructor(props: ISocialMediaProps, state: ISocialMediaState) {
    super(props);
    this.state = {
      allVideos: [],
    };
  }
  public render(): React.ReactElement<ISocialMediaProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <section>
        <div className="Stay_Connected">
          <h2 className="Stay_Connected_title">Social Media</h2>
          <div className="p-15">
            {/* <iframe style={{ border: "none" }} height={"420px"} width={"100%"} data-tweet-url="https://www.linkedin.com/company/netuno-usa-inc" src="data:text/html;charset=utf-8,%3Ca%20class%3D%22twitter-timeline%22%20href%3D%22https%3A//twitter.com/iZOOlogic%3Fref_src%3Dtwsrc%255Etfw%22%3ETweets%20by%20iZOOlogic%3C/a%3E%0A%3Cscript%20async%20src%3D%22https%3A//platform.twitter.com/widgets.js%22%20charset%3D%22utf-8%22%3E%3C/script%3E%0A"></iframe> */}
            <div className="ms-Grid">
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <InstagramEmbed />
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <div className="ms-Grid-row">
                    {this.state.allVideos.length > 0 &&
                      this.state.allVideos.slice(0, 2).map((ele, ind) => (
                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6" key={ele.FileLeafRef}>
                          <video controls width="100%">
                            <source src={`https://bullioninternational.sharepoint.com/${ele.FileRef}`} type="video/mp4" />
                          </video>
                        </div>
                      ))}
                    {this.state.allVideos[2] && (
                      <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" key={this.state.allVideos[2].FileLeafRef}>
                        <video controls width="100%">
                          <source src={`https://bullioninternational.sharepoint.com/${this.state.allVideos[2].FileRef}`} type="video/mp4" />
                        </video>
                      </div>
                    )}
                    {/* <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
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
                    </div> */}
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

  public componentDidMount = async () => {
    this.loadVideos();
  };

  private loadVideos = async () => {
    try {
      // Change "Videos" to your library name
      const items = await sp.web.lists
        .getByTitle("Videos")
        .items.select("FileRef", "FileLeafRef")
        .top(3)()
        .then((res) => res.filter((i) => i.FileLeafRef.endsWith(".mp4")));

      // setVideos(items);
      this.setState({ allVideos: items });
      console.log(items);
    } catch (error) {
      console.error("Error loading videos:", error);
    }
  };
}
