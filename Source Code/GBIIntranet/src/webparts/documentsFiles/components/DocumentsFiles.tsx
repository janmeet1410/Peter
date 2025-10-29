import * as React from "react";
import styles from "./DocumentsFiles.module.scss";
import { IDocumentsFilesProps } from "./IDocumentsFilesProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import { Icon } from "office-ui-fabric-react";

require('../assets/style.css')
export interface IDocumentsFilesState {
  departmentFiles: any;
}

export default class DocumentsFiles extends React.Component<IDocumentsFilesProps, IDocumentsFilesState> {
  constructor(props: IDocumentsFilesProps, state: IDocumentsFilesState) {
    super(props);
    this.state = {
      departmentFiles: [],
    };
  }
  public render(): React.ReactElement<IDocumentsFilesProps> {
    const { description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName } = this.props;

    return (
      <div>
        {/* <h3>Benefits Files</h3> */}
        <div className="newswrapper">
          {this.state.departmentFiles.length > 0 &&
            this.state.departmentFiles.map((file, ind) => {
              return (
                <a href={file.Url} target="_blank" data-interception="off" className="file-item" key={ind}>
                  {/* Bind image only if item is "url type" (link) */}
                  {file.IsLinkType && file.ImageUrl ? <img src={file.ImageUrl} alt={file.Name} style={{ maxWidth: 90, maxHeight: 40, marginRight: 8 }} /> : <Icon iconName="FabricFolder" />}
                  <div>{file.Name}</div>
                </a>
              );
            })}
        </div>
      </div>
    );
  }

  public componentDidMount = async () => {
    await this.getDepartmentFiles();
  };

  private getDepartmentFiles = async () => {
    try {
      const folderServerRelativeUrl = `${this.props.serverrelativeUrl}/Benefits`;
      const allItems = await sp.web.lists
        .getByTitle("Benefits")
        .items.select(
          "ID",
          "File/Name",
          "File/ServerRelativeUrl",
          "File/TimeCreated",
          "File/TimeLastModified",
          "File/Author/Title",
          "Folder/Name",
          "Folder/ServerRelativeUrl",
          "FileDirRef",
          "Image" // <-- Image column
        )
        .expand("File", "Folder", "File/Author")
        // .filter('ID eq 12')
        .top(5000)
        .get();

      // Only items in root directory (no folders, no subfolders)
      const mainItems = allItems.filter((item) => item.FileDirRef === folderServerRelativeUrl);

      const mappedItems = mainItems.map((item) => {
        const isFolder = !!item.Folder;
        const name = isFolder ? item.Folder.Name : item.File.Name;
        const url = isFolder ? item.Folder.ServerRelativeUrl : item.File.ServerRelativeUrl + "?web=1";
        const created = isFolder ? null : item.File.TimeCreated;
        const modified = isFolder ? null : item.File.TimeLastModified;
        const author = isFolder ? null : item.File.Author?.Title || "";
        // Only bind if link type (you may use alternative logic, e.g., custom column)
        const isLinkType = !isFolder && item.File.Name.toLowerCase().endsWith(".url");
        // Modern SharePoint Image column returns an object. Use ?. syntax in case null
        let imgtemp  = item.Image ? JSON.parse(item.Image) : null
    const serverRelativeUrl = imgtemp ? imgtemp.serverRelativeUrl : null;
        const imageUrl =
          item.Image 
            ? serverRelativeUrl || item.Image.Url // Try both keys for compatibility
            : null;
        // Fallback image if not found
        const fallback = require("../assets/iconDoc.png");
        return {
          Type: isFolder ? "Folder" : "File",
          Name: name,
          Url: url,
          Created: created,
          Modified: modified,
          Author: author,
          IsLinkType: isLinkType,
          ImageUrl: isLinkType && imageUrl ? imageUrl : fallback, // Only bind image if link type
        };
      });

      const sortedItems = mappedItems.sort((a, b) => a.Name.toLowerCase().localeCompare(b.Name.toLowerCase()));
      this.setState({ departmentFiles: sortedItems });
    } catch (error) {
      console.error("Error fetching files with department:", error);
    }
  };
}
