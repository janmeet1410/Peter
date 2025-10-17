import * as React from "react";
import styles from "./Poilicies.module.scss";
import { IPoiliciesProps } from "./IPoiliciesProps";
import { escape } from "@microsoft/sp-lodash-subset";

require("../assets/style.css");

export default class Poilicies extends React.Component<IPoiliciesProps, {}> {
  public render(): React.ReactElement<IPoiliciesProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <section>
        <div className="policies-container">
          <h2>Policies</h2>
          <div className="policy-card">
            <div className="policy-header">
              <span className="policy-icon">📄</span>
              <h3>📄 Privacy & Policy</h3>
              <a href="#" className="policy-link">
                ↗
              </a>
            </div>
            <p>This Privacy Policy applies to information about you that Gold Bullion International, LLC (“GBI”, “we”, “our”, “us”) may obtain when you visit or interact with BullionInternational.com or direct.BullionInternational.com (collectively, the “Sites”). This Privacy Policy describes how we may collect, use, or share information ab...</p>
          </div>

          <div className="policy-card">
            <div className="policy-header">
              <span className="policy-icon">📋</span>
              <h3>Collection Of Personal Information</h3>
              <a href="#" className="policy-link">
                ↗
              </a>
            </div>
            <p>Information you provide: We may obtain information that you provide directly to us on the Sites (e.g., when you contact us, create an account, or make a purchase). This information may include personal information, which is information that can be used to identify you individually, such as your name, email address, postal ad...</p>
          </div>

          <div className="policy-card">
            <div className="policy-header">
              <span className="policy-icon">💡</span>
              <h3>Use Of Personal Information</h3>
              <a href="#" className="policy-link">
                ↗
              </a>
            </div>
            <p>We may use personal information we collect through the Sites f...</p>
            <ul>
              <li>Facilitating and personalizing your user experience;</li>
            </ul>
          </div>

          <div className="policy-card">
            <div className="policy-header">
              <span className="policy-icon">👥</span>
              <h3>Collection Of Personal Information</h3>
              <a href="#" className="policy-link">
                ↗
              </a>
            </div>
            <p>Information you provide: We may obtain information that you provide directly to us on the Sites (e.g., when you contact us, create an account, or make a purchase). This information may include personal information, which is information that can be used to identify you individually, such as your name, email address, postal ad...</p>
          </div>
        </div>
      </section>
    );
  }
}
