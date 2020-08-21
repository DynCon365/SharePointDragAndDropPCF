/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";
import { useDropzone } from "react-dropzone";
import { Stack } from "office-ui-fabric-react/lib/Stack";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { initializeIcons } from "@uifabric/icons";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import * as Msal from "msal";
import { sp } from "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";
import { PnPFetchClient } from "./PnPFetchClient";
import { getSharePointFolderName } from "./Helper";
import { Guid } from "./Guid";

initializeIcons();

export interface ISharePointUploadProperties {
  entityPrimaryFieldLogicalName: string;
  useRelationship: boolean;
  relationshipLogicalName: string;
  relationshipPrimaryFieldLogicalName: string;
  controlToRefresh: string;
  clientId: string;
  loginHint: string;
  sharePointSiteId: string;
  context: ComponentFramework.Context<IInputs>;
}

const iconClass = mergeStyles({
  fontSize: 50,
  height: 50,
  width: 50,
  margin: "0 auto",
});

export const SharePointAttachmentUploadControl: React.FC<ISharePointUploadProperties> = (
  sharePointUploadProps: ISharePointUploadProperties,
) => {
  console.log("Control Loaded");
  const context = sharePointUploadProps.context;
  const useRelationship = context.parameters.useRelationship.raw === "true";
  const sharePointSiteGuid: string = context.parameters.sharePointSiteGuid.raw ?? "";
  //@ts-ignore
  const entityId: string = context.mode.contextInfo.entityId;
  //@ts-ignore
  const entityName: string = context.mode.contextInfo.entityTypeName;

  if (entityId == null || entityId === "" || !Guid.isValid(entityId) || Guid.isEmpty(entityId)) {
    return (
      <Stack>
        <Icon iconName="ErrorBadge" className={iconClass} />
        <p>Please save record to enable control.</p>
      </Stack>
    );
  }
  let sharePointSite: ComponentFramework.WebApi.Entity;
  let primaryEntity: ComponentFramework.WebApi.Entity;
  let parentEntity: ComponentFramework.WebApi.Entity | undefined;
  let currentUser: ComponentFramework.WebApi.Entity;
  const [totalFileCount, setTotalFileCount] = React.useState(0);
  const [currentUploadCount, setCurrentUploadCount] = React.useState(0);
  const onDrop = React.useCallback(
    (acceptedFiles: any) => {
      console.log("On Drop Triggered");
      if (acceptedFiles && acceptedFiles.length) {
        setTotalFileCount(acceptedFiles.length);
      }
      const getSharePointSite = async function () {
        try {
          console.log("Getting SharePoint Site");
          const response = await context.webAPI.retrieveRecord("sharepointsite", sharePointSiteGuid);
          sharePointSite = response;
          return response;
        } catch (error) {
          const errOptions = { message: error.message };
          context.navigation.openErrorDialog(errOptions);
        }
      };
      const getPrimaryEntity = async (entityName: string, entityId: string) => {
        try {
          console.log("Getting Primary Entity");
          const response = await context.webAPI.retrieveRecord(entityName, entityId);
          primaryEntity = response;
          return response;
        } catch (error) {
          const errOptions = { message: error.message };
          context.navigation.openErrorDialog(errOptions);
        }
      };
      const getParentEntity = async (parentEntityTypeName: string) => {
        try {
          if (
            useRelationship &&
            primaryEntity !== undefined &&
            primaryEntity["_" + sharePointUploadProps.relationshipLogicalName + "_value"] !== undefined &&
            primaryEntity["_" + sharePointUploadProps.relationshipLogicalName + "_value"] !== ""
          ) {
            console.log("Getting Parent Entity");
            const relationshipId = primaryEntity["_" + sharePointUploadProps.relationshipLogicalName + "_value"];
            const response = await context.webAPI.retrieveRecord(parentEntityTypeName, relationshipId);
            parentEntity = response;
            return response;
          } else {
            return undefined;
          }
        } catch (error) {
          const errOptions = { message: error.message };
          context.navigation.openErrorDialog(errOptions);
        }
      };
      const getCurrentUser = async () => {
        try {
          console.log("Getting User");
          if (context.userSettings.userName.length) {
            const response = await context.webAPI.retrieveRecord("systemuser", context.userSettings.userId);
            currentUser = response;
            return response;
          } else {
            const errOptions = { message: "User not found in Context." };
            context.navigation.openErrorDialog(errOptions);
          }
        } catch (error) {
          const errOptions = { message: error.message };
          context.navigation.openErrorDialog(errOptions);
        }
      };

      const processAcceptedFiles = async (files: any) => {
        try {
          await getSharePointSite();
          let sharePointFolderName: string;
          const sharePointAboluteUrl: string = sharePointSite["absoluteurl"];
          const sharePointRelativeUrl: string = sharePointAboluteUrl.replace(/^(?:\/\/|[^\/]+)*\//, "");
          const sharePointStructureEntity: string = sharePointSite["folderstructureentity"];
          let primaryEntityFolderUrl: string;
          let parentEntityFolderUrl: string;

          await getPrimaryEntity(entityName, entityId);
          await getParentEntity(sharePointStructureEntity);
          await getCurrentUser();

          const msalConfig = {
            auth: {
              clientId: sharePointUploadProps.clientId,
            },
          };

          const ssoRequest = {
            loginHint: currentUser["domainname"],
          };
          const msalInstance = new Msal.UserAgentApplication(msalConfig);
          msalInstance
            .ssoSilent(ssoRequest)
            .then((response) => {
              sp.setup({
                sp: {
                  fetchClientFactory: () => {
                    return new PnPFetchClient(msalInstance);
                  },
                },
              });
            })
            .catch((error) => {
              const errOptions = { message: `${error.message}` };
              context.navigation.openErrorDialog(errOptions);
            });
          const web = Web(sharePointAboluteUrl);

          const sharepointFolderExists = async (folderName: string) => {
            try {
              console.log("Checking if SharePointFolder Exists");
              const folderExists: boolean = (await web.getFolderByServerRelativePath(folderName)()).Exists;
              console.log("SharePoint Folder Exists");
              return folderExists;
            } catch {
              console.log("SharePoint Folder Not Found - Will be Created");
              return false;
            }
          };
          const createSharePointFolder = async (parentFolder: string, newFolder: string) => {
            await web.getFolderByServerRelativePath(parentFolder).folders.add(newFolder);
          };

          for (let i = 0; i < acceptedFiles.length; i++) {
            setCurrentUploadCount(i);
            console.log("Looping Accepted Files and Uploading to SharePoint");

            //All cases: Set folder relative URL & get Primary Entity Folder Url
            sharePointFolderName = "/" + sharePointRelativeUrl + "/";

            primaryEntityFolderUrl =
              getSharePointFolderName(primaryEntity[sharePointUploadProps.entityPrimaryFieldLogicalName]) +
              "_" +
              entityId.replace(/-/g, "").toUpperCase();

            //CASE 1: SharePoint Set to No Primary Entity Structure
            //Example: /sites/Dynamics/lead/Fabrikam_Guid
            if (
              sharePointStructureEntity.toLocaleLowerCase() === "none" ||
              sharePointStructureEntity === undefined ||
              sharePointStructureEntity === "" ||
              sharePointStructureEntity.length === 0
            ) {
              console.log("Using SharePoint Type: 1");
              sharePointFolderName += entityName + "/" + primaryEntityFolderUrl;
              const folderResult = await sharepointFolderExists(sharePointFolderName);
              if (!folderResult) {
                await createSharePointFolder("/" + sharePointRelativeUrl + "/" + entityName, primaryEntityFolderUrl);
              }
            } else {
              //CASE 2: SharePoint Set to Primary Entity Structure & Use Relationship
              //Example /sites/Dynamics/account/Fabrikam_Guid/opportunity/OpportunityName_Guid
              if (
                useRelationship &&
                parentEntity !== undefined &&
                parentEntity[sharePointStructureEntity + "id"] !== undefined
              ) {
                console.log("Using SharePoint Type: 2");
                parentEntityFolderUrl = getSharePointFolderName(
                  parentEntity[sharePointUploadProps.relationshipPrimaryFieldLogicalName] +
                    "_" +
                    parentEntity[sharePointStructureEntity + "id"].replace(/-/g, "").toUpperCase(),
                );
                sharePointFolderName +=
                  sharePointStructureEntity +
                  "/" +
                  parentEntityFolderUrl +
                  "/" +
                  entityName +
                  "/" +
                  primaryEntityFolderUrl;

                // //Check Folder Tree
                let folderTreeParent = "/" + sharePointRelativeUrl + "/" + sharePointStructureEntity;
                let folderResult = await sharepointFolderExists(folderTreeParent + "/" + parentEntityFolderUrl);
                if (!folderResult) {
                  console.log("Creating folder: " + folderTreeParent + "/" + parentEntityFolderUrl);
                  await createSharePointFolder(folderTreeParent, parentEntityFolderUrl);
                }
                folderTreeParent += "/" + parentEntityFolderUrl;
                folderResult = await sharepointFolderExists(folderTreeParent + "/" + entityName);
                if (!folderResult) {
                  console.log("Creating folder: " + folderTreeParent + "/" + entityName);
                  await createSharePointFolder(folderTreeParent, entityName);
                }
                folderTreeParent += "/" + entityName;
                folderResult = await sharepointFolderExists(folderTreeParent + "/" + primaryEntityFolderUrl);
                if (!folderResult) {
                  console.log("Creating folder: " + folderTreeParent + "/" + primaryEntityFolderUrl);
                  await createSharePointFolder(folderTreeParent, primaryEntityFolderUrl);
                }
              } else {
                //CASE 3: SharePoint Set to Primary Entity Structure but No Relationship for Upload
                //Example /sites/Dynamics/account/Fabrikam_Guid
                console.log("Using SharePoint Type: 3");
                sharePointFolderName += entityName + "/" + primaryEntityFolderUrl;
                const folderResult = await sharepointFolderExists(sharePointFolderName);
                if (!folderResult) {
                  console.log(
                    "Creating folder: " +
                      sharePointRelativeUrl +
                      "/" +
                      sharePointStructureEntity +
                      "/" +
                      primaryEntityFolderUrl,
                  );
                  await createSharePointFolder(
                    "/" + sharePointRelativeUrl + "/" + sharePointStructureEntity,
                    primaryEntityFolderUrl,
                  );
                }
              }
            }

            const file = acceptedFiles[i] as any;

            if (file.size <= 10485760) {
              await web.getFolderByServerRelativeUrl(sharePointFolderName).files.add(file.name, file, true);
            } else {
              await web.getFolderByServerRelativeUrl(sharePointFolderName).files.addChunked(
                file.name,
                file,
                (data) => {
                  console.log(data);
                },
                true,
              );
            }
          }
          console.log("File Processing Complete");
          setTotalFileCount(0);
        } catch (e) {
          const errorMessagePrefix =
            acceptedFiles.length === 1
              ? "An error has occurred while trying to upload the attachment."
              : "One or more errors occured when trying to upload the attachments.";
          const errOptions = { message: `${errorMessagePrefix} ${e.message}` };
          context.navigation.openErrorDialog(errOptions);
        }
      };
      processAcceptedFiles(acceptedFiles);
    },
    [totalFileCount, currentUploadCount],
  );
  console.log("Setting Dropzone Properties");
  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop });

  let fileUploadStatus = null;
  if (totalFileCount > 0) {
    fileUploadStatus = (
      <div className={"filesStatsCont uploadDivs"}>
        <div className={"uploadStatusText"}>
          Uploading ({currentUploadCount}/{totalFileCount})
        </div>
      </div>
    );
  }
  return (
    <Stack>
      <div {...getRootProps()} style={{ backgroundColor: isDragActive ? "#F8F8F8" : "white" }}>
        <input {...getInputProps()} />
        <Icon iconName="CloudUpload" className={iconClass} />
        <p>Drag &amp; Drop Some Files Here!</p>
      </div>
      {fileUploadStatus}
    </Stack>
  );
};
