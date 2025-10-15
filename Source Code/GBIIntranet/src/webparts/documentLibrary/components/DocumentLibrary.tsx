import * as React from 'react';
import styles from './DocumentLibrary.module.scss';
import { IDocumentLibraryProps } from './IDocumentLibraryProps';
import { escape } from '@microsoft/sp-lodash-subset';

require('../assets/style.css');

export default class DocumentLibrary extends React.Component<IDocumentLibraryProps, {}> {
  public render(): React.ReactElement<IDocumentLibraryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
<section className="library-section">
    <h2>Document library</h2>

    <div className="library-cards">

      <div className="library-card" >
        <div className="overlay"></div>
        <div className="content">
          <h3>Library Facilities</h3>
          <ul>
            <li>⭐ Reading Halls</li>
            <li>⭐ Magazine & Periodical Section</li>
            <li>⭐ Cafeteria</li>
          </ul>
        </div>
      </div>

      <div className="library-card" >
        <div className="overlay"></div>
        <div className="content">
          <h3>Library Services</h3>
          <ul>
            <li>⭐ Book Reservation & Renewals</li>
            <li>⭐ New Arrivals Section</li>
            <li>⭐ Research & Career Guidance</li>
          </ul>
        </div>
      </div>

      <div className="library-card" >
        <div className="overlay"></div>
        <div className="content">
          <h3>Digital Library</h3>
          <ul>
            <li>⭐ E-Books & PDFs</li>
            <li>⭐ Digital Access & App</li>
            <li>⭐ Online Databases</li>
          </ul>
        </div>
      </div>

    </div>
  </section>
    );
  }
}
