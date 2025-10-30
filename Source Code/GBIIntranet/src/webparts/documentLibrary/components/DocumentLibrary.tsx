import * as React from "react";
import styles from "./DocumentLibrary.module.scss";
import { IDocumentLibraryProps } from "./IDocumentLibraryProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";

require("../assets/style.css");
export interface IDocumentLibraryState {
  resourcesFiles: any;
}
export default class DocumentLibrary extends React.Component<IDocumentLibraryProps, IDocumentLibraryState> {
  constructor(props: IDocumentLibraryProps, state: IDocumentLibraryState) {
    super(props);
    this.state = {
      resourcesFiles: [],
    };
  }
  public render(): React.ReactElement<IDocumentLibraryProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <section className="library-section">
        <h2>Document library</h2>

        <div className="library-cards">
          {this.state.resourcesFiles.length > 0 &&
            this.state.resourcesFiles.map((file, ind) => {
              return (
                <a href={file.Url} className="library-card">
                  <div className="overlay"></div>
                  <div className="content1">
                    <h3>{file.Name}</h3>
                    <ul>
                      {file.DocumentTag.length > 0 &&
                        file.DocumentTag.map((el, ind) => {
                          return <li>⭐ {el}</li>;
                        })}
                    </ul>
                  </div>
                </a>
              );
            })}
          {/* <div className="library-card">
            <div className="overlay"></div>
            <div className="content1">
              <h3>Library Facilities</h3>
              <ul>
                <li>⭐ Reading Halls</li>
                <li>⭐ Magazine & Periodical Section</li>
                <li>⭐ Cafeteria</li>
              </ul>
            </div>
          </div>

          <div className="library-card">
            <div className="overlay"></div>
            <div className="content1">
              <h3>Library Services</h3>
              <ul>
                <li>⭐ Book Reservation & Renewals</li>
                <li>⭐ New Arrivals Section</li>
                <li>⭐ Research & Career Guidance</li>
              </ul>
            </div>
          </div>

          <div className="library-card">
            <div className="overlay"></div>
            <div className="content1">
              <h3>Digital Library</h3>
              <ul>
                <li>⭐ E-Books & PDFs</li>
                <li>⭐ Digital Access & App</li>
                <li>⭐ Online Databases</li>
              </ul>
            </div>
          </div> */}
        </div>
      </section>
    );
  }

  public componentDidMount = async () => {
    await this.getResourcesFiles();
  }

  private getResourcesFiles = async () => {
    try {
      const folderServerRelativeUrl = `${this.props.serverrelativeUrl}/Shared Documents`;
      const allItems = await sp.web.lists.getByTitle("Documents").items.select("ID", "File/Name", "File/ServerRelativeUrl", "File/TimeCreated", "File/TimeLastModified", "File/Author/Title", "Folder/Name", "Folder/ServerRelativeUrl", "Tags", "FileDirRef").expand("File", "Folder", "File/Author").top(5000).get();

      // Only items under the specified folder (direct children only)
      // const items = allItems.filter((item) => item.FileDirRef === folderServerRelativeUrl.replace(/^\//, ""));
      // Only items in root directory (no folders, no subfolders)
      const mainItems = allItems.filter((item) => item.FileDirRef === folderServerRelativeUrl);

      const mappedItems = mainItems.map((item) => {
        const isFolder = !!item.Folder;
        const name = isFolder ? item.Folder.Name : item.File.Name;
        const url = isFolder ? item.Folder.ServerRelativeUrl : item.File.ServerRelativeUrl + "?web=1";
        const created = isFolder ? null : item.File.TimeCreated;
        const modified = isFolder ? null : item.File.TimeLastModified;
        const author = isFolder ? null : item.File.Author?.Title || "";
        const documenttag = item.Tags ? item.Tags.split(",").map((item) => item.trim()) : "";
        return {
          Type: isFolder ? "Folder" : "File",
          Name: name,
          Url: url,
          Created: created,
          Modified: modified,
          Author: author,
          DocumentTag: documenttag,
        };
      });

      const sortedItems = mappedItems.sort((a, b) => a.Name.toLowerCase().localeCompare(b.Name.toLowerCase()));
      this.setState({ resourcesFiles: sortedItems });
    } catch (error) {
      console.error("Error fetching root resources with department:", error);
    }
  }
}
