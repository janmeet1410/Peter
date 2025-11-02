import * as React from "react";
import styles from "./NewsAndAnnouncements.module.scss";
import { INewsAndAnnouncementsProps } from "./INewsAndAnnouncementsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";

require("../assets/CSS/style.css");

export interface INewsAndAnnouncementsState {
  CompanyNewsData: any;
  BreakingNewsData: any;
}

export default class NewsAndAnnouncements extends React.Component<INewsAndAnnouncementsProps, INewsAndAnnouncementsState> {
  constructor(props: INewsAndAnnouncementsProps, state: INewsAndAnnouncementsState) {
    super(props);
    this.state = {
      CompanyNewsData: [],
      BreakingNewsData: [],
    };
  }
  public render(): React.ReactElement<INewsAndAnnouncementsProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    function SampleNextArrow(props) {
      const { className, style, onClick } = props;
      return <img className={className + " arrow-img-icon"} src={require("../assets/Icons/RightArrow.svg")} style={{ ...style }} onClick={onClick} />;
    }

    function SamplePrevArrow(props) {
      const { className, style, onClick } = props;
      return <img className={className + " arrow-img-icon"} src={require("../assets/Icons/LeftArrow.svg")} style={{ ...style }} onClick={onClick} />;
    }

    var settings = {
      dots: true,
      infinite: true,
      speed: 1000,
      slidesToShow: 3,
      slidesToScroll: 1,
      autoplay: false,
      responsive: [
        {
          breakpoint: 720,
          settings: {
            slidesToShow: 1,
          },
        },
      ],
      nextArrow: <SampleNextArrow />,
      prevArrow: <SamplePrevArrow />,
    };

    return (
      <section >
        <div className="container">
          {/* <div className='BreakingNewContainer'>
            <div><img src={require("../assets/Icons/BreakingNews.png")} /></div>
            <div><a style={{ textDecoration: "none", cursor: "pointer" }} href={this.state.BreakingNewsData.length > 0 ? this.state.BreakingNewsData[0]?.URL ? this.state.BreakingNewsData[0]?.URL.Url : "#" : "#"}><p><span><b style={{ color: "#113F8A" }}>News</b></span> {this.state.BreakingNewsData.length > 0 ? this.state.BreakingNewsData[0]?.Description : ""}</p></a></div>
          </div> */}
          <div>
            <h2 className="News-Header">News & Announcements</h2>
          </div>
          <div>
            <Slider {...settings}>
              {this.state.CompanyNewsData &&
                this.state.CompanyNewsData.map((news, nwIdx) => {
                  let imageURL = news.AttachmentFiles.length > 0 ? news.AttachmentFiles[0].ServerRelativeUrl : news.Image ? JSON.parse(news.Image).serverRelativeUrl : require(`../assets/Icons/BreakingNews.png`);
                  return (
                    <a  href={news.URL ? news.URL.Url : "#"} className="newsMain">
                      <div>
                        <img src={imageURL} />
                      </div>
                      <div>
                        <a style={{ textDecoration: "none", cursor: "pointer" }} href={news.URL ? news.URL.Url : "#"}>
                          <p className="New-description">{news.Description}</p>
                        </a>
                      </div>
                    </a>
                  );
                })}
            </Slider>
          </div>
        </div>
      </section>
    );
  }

  public componentDidMount = async () => {
    await this.getCompanyNewsDetails();
  };

  private getCompanyNewsDetails = async () => {
    try {
      const CompanyNewsDetails = await sp.web.lists.getByTitle("Announcements").items.select("ID,Title,URL,Description,Image").expand("AttachmentFiles").orderBy("Modified", false).getAll();

      if (CompanyNewsDetails.length > 0) {
        this.setState({ CompanyNewsData: CompanyNewsDetails, BreakingNewsData: [CompanyNewsDetails[0]] });
      }
    } catch (error) {
      console.log(error);
    }
  };
}
