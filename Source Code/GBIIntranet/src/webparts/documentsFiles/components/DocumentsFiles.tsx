import * as React from "react";
import styles from "./DocumentsFiles.module.scss";
import { IDocumentsFilesProps } from "./IDocumentsFilesProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import { Icon } from "office-ui-fabric-react";

require("../assets/style.css");
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
      const allItems = await sp.web.lists.getByTitle("Benefits").items.select("ID", "File/Name", "File/ServerRelativeUrl", "File/TimeCreated", "File/TimeLastModified", "File/Author/Title", "Folder/Name", "Folder/ServerRelativeUrl", "FileDirRef", "Image").expand("File", "Folder", "File/Author").top(5000).get();

      // Only items in root directory (no folders, no subfolders)
      const mainItems = allItems.filter((item) => item.FileDirRef === folderServerRelativeUrl);

      // Map items and resolve .url links
      const mappedItems = await Promise.all(
        mainItems.map(async (item) => {
          const isFolder = !!item.Folder;
          const name = isFolder ? item.Folder.Name : item.File.Name;
          let url = isFolder ? item.Folder.ServerRelativeUrl : item.File.ServerRelativeUrl + "?web=1";
          const created = isFolder ? null : item.File.TimeCreated;
          const modified = isFolder ? null : item.File.TimeLastModified;
          const author = isFolder ? null : item.File.Author?.Title || "";
          const isLinkType = !isFolder && item.File.Name.toLowerCase().endsWith(".url");

          // Resolve URL for .url files
          if (isLinkType) {
            try {
              const fileContent = await sp.web.getFileByServerRelativeUrl(item.File.ServerRelativeUrl).getText();
              const match = fileContent.match(/URL=(.*)/);
              if (match && match[1]) {
                url = match[1].trim();
              }
            } catch (err) {
              console.error("Failed to fetch or parse .url file content:", err);
              // fallback stays as SharePoint file url
            }
          }

          let imgtemp = item.Image ? JSON.parse(item.Image) : null;
          const serverRelativeUrl = imgtemp ? imgtemp.serverRelativeUrl : null;
          const imageUrl = item.Image ? serverRelativeUrl || item.Image.Url : null;
          const fallback = require("../assets/iconDoc.png");

          return {
            Type: isFolder ? "Folder" : "File",
            Name: name,
            Url: url,
            Created: created,
            Modified: modified,
            Author: author,
            IsLinkType: isLinkType,
            ImageUrl: isLinkType && imageUrl ? imageUrl : fallback,
          };
        })
      );

      // Sort by priority: Link type first, then files, then folders, then alphabetically by name
      const sortedItems = mappedItems.sort((a, b) => {
        const getPriority = (item) => {
          if (item.IsLinkType) return 0;
          if (item.Type === "File") return 1;
          if (item.Type === "Folder") return 2;
          return 3;
        };

        const priorityA = getPriority(a);
        const priorityB = getPriority(b);

        if (priorityA !== priorityB) {
          return priorityA - priorityB;
        }
        return a.Name.toLowerCase().localeCompare(b.Name.toLowerCase());
      });

      this.setState({ departmentFiles: sortedItems });
    } catch (error) {
      console.error("Error fetching files with department:", error);
    }
  };
}
