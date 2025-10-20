import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';

require('../assets/css/fabric.min.css');

import {
  Version,
  Environment,
  EnvironmentType,
  ServiceScope,
  Log,
  Text,
} from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Persona, PersonaInitialsColor, PersonaSize, IPersonaProps } from '@fluentui/react/lib/Persona';
import { DocumentCard, DocumentCardType, Icon } from '@fluentui/react';
import { Pagination } from '../components/Pagination';


const EXP_SOURCE: string = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID: string =
  '914330ee-2df2-4f6e-a858-30c23a812408';

export class PersonaCard extends React.Component<IPersonaCardProps, IPersonaCardState> {
  constructor(props: IPersonaCardProps) {
    super(props);

    this.state = {
      livePersonaCard: undefined,
      pictureUrl: undefined,
    };
  }


  public async componentDidMount() {

  }

  public componentDidUpdate(
    prevProps: IPersonaCardProps,
    prevState: IPersonaCardState
  ): void { }




  public render(): React.ReactElement<IPersonaCardProps> {
    {
      return (
        <div className={'ms-Grid-col ms-sm3 gridSm4'}>
          {/* <div className={styles.personaContainer} onClick={() => window.open("https://delve.office.com/?u=" + this.props.profileProperties.DelveUrl, "_blank")}> */}

          {/* <Persona
                  text={this.props.profileProperties.DisplayName}
                  // secondaryText={"Job Title: " + (this.props.profileProperties.Department ? this.props.profileProperties.JobTitle : "")}
                  // tertiaryText={"Department: " + (this.props.profileProperties.Department ? this.props.profileProperties.Department : "")}
                  imageUrl={this.props.profileProperties.ImgUrl}
                  size={PersonaSize.size72}
                  imageShouldFadeIn={false}
                  imageShouldStartVisible={true}
                >

                  <div>
                    <span className="spnUserDetails" title={this.props.profileProperties.JobTitle}>
                      {"Job Title: " + (this.props.profileProperties.JobTitle ? this.props.profileProperties.JobTitle : "")}
                    </span>
                  </div>

                  <div>
                    <span className="spnUserDetails" title={this.props.profileProperties.Department}>
                      {"Department: " + (this.props.profileProperties.Department ? this.props.profileProperties.Department : "")}
                    </span>
                  </div>

                  {this.props.profileProperties.Email ? (
                    <div>
                      <span className="spnUserDetails" title={this.props.profileProperties.Email}>
                        {"Email: " + this.props.profileProperties.Email}
                      </span>
                    </div>
                  ) : (
                    ''
                  )}

                  <div>
                    <span className="spnUserDetails" title={this.props.profileProperties.OfficeLocation}>
                      {"Office Location: " + (this.props.profileProperties.OfficeLocation ? this.props.profileProperties.OfficeLocation : "")}
                    </span>
                  </div>

                  <div>
                    <Icon iconName="TeamsLogo" title="Teams" className="iconCursorPointer" onClick={() => window.open("https://teams.microsoft.com/l/chat/0/0?users=" + this.props.profileProperties.Email)} />
                    <Icon iconName="DelveLogo" title="Teams" className="iconCursorPointer"  />
                  </div>
                </Persona> */}
          <div>
            <div className='profile-card'>
              {/* <div><img src={require("../../webparts/companyDirectory/assets/user-profile.jpg")} alt="" /></div> */}
              <div><img src={this.props.profileProperties.ImgUrl} alt="" /></div>
              <div className='deatails-wrapper'>
                <p className='user-name' title={this.props.profileProperties.DisplayName}>{this.props.profileProperties.DisplayName}</p>
                <p className='job-title' title={this.props.profileProperties.JobTitle}> {(this.props.profileProperties.JobTitle ? this.props.profileProperties.JobTitle : "--")}</p>
                <p className='Department' title={this.props.profileProperties.Department}>{(this.props.profileProperties.Department ? this.props.profileProperties.Department : "--")}</p>
                <p className='cell-phone' title={this.props.profileProperties.CellPhone}>{(this.props.profileProperties.CellPhone ? this.props.profileProperties.CellPhone : "--")}</p>
                <p className='location' title={this.props.profileProperties.OfficeLocation}>{(this.props.profileProperties.OfficeLocation ? this.props.profileProperties.OfficeLocation : "--")}</p>
                <p className='email' title={this.props.profileProperties.Email}> {this.props.profileProperties.Email ? this.props.profileProperties.Email : "--"}</p>
              </div>
              <div className='User-quicklinks'>
                <Icon iconName="TeamsLogo" title="Teams" className="iconCursorPointer" onClick={() => window.open("https://teams.microsoft.com/l/chat/0/0?users=" + this.props.profileProperties.Email)} />
                <Icon iconName="OutlookLogo" title="Outlook" className="iconCursorPointer" onClick={() => window.open("mailto:" + this.props.profileProperties.Email, "_blank")} />
              </div>
            </div>

          </div>
        </div>
      );
    }
  }
}
