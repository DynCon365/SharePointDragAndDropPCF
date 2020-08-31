/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { ServiceProvider } from "../components/pcf-react/ServiceProvider";
import { IInputs } from "../generated/ManifestTypes";
import { ControlContextService } from "../components/pcf-react/ControlContextService";
import { ParametersChangedEventArgs } from "../components/pcf-react/ParametersChangedEventArgs";
import { decorate, observable, action } from "mobx";
import { CdsService } from "./CdsService";
import { SharePointService } from "./SharePointService";
import { containsInvalidFileFolderChars } from "@pnp/sp";

export class SharePointUploadControlVM {
  serviceProvider: ServiceProvider;
  controlContext: ControlContextService;
  cdsService: CdsService;
  sharePointService: SharePointService;
  sharePointRelativeUrl: string;
  isLoading: boolean;
  isEnabled: boolean;
  droppedFiles: number;
  currentFile: number;
  filesProcessed: number;
  isProcessingFiles: boolean;

  constructor(serviceProvider: ServiceProvider) {
    this.serviceProvider = serviceProvider;
    this.controlContext = serviceProvider.get<ControlContextService>(ControlContextService.serviceProviderName);
    this.controlContext.onLoadEvent.subscribe(this.onLoad);
    this.controlContext.onParametersChangedEvent.subscribe(this.onInParametersChanged);
  }

  async onLoad(): Promise<void> {
    console.log("onLoad");
    this.isLoading = true;
    this.isEnabled = true;
    this.cdsService = this.serviceProvider.get<CdsService>("CdsService");
    this.sharePointService = this.serviceProvider.get<SharePointService>("SharePointService");
    await this.getCdsRecords();
    this.droppedFiles = 0;
  }

  onInParametersChanged(context: ControlContextService, args: ParametersChangedEventArgs): void {
    console.log("onInParametersChanged");
  }

  async getCdsRecords(): Promise<void> {
    await this.cdsService.getSharePointSite();
    await this.cdsService.getPrimaryEntity();
    await this.cdsService.getParentEntity();
    await this.cdsService.getCurrentUser();
    await this.sharePointService.setupSharePoint(
      this.cdsService.currentUser["domainname"],
      this.cdsService.sharepointSite["absoluteurl"],
      this.cdsService.sharepointSite["folderstructureentity"],
    );
    this.sharePointRelativeUrl = this.cdsService.sharepointSite["absoluteurl"].replace(/^(?:\/\/|[^\/]+)*\//, "");
    this.isLoading = false;
  }

  async onFileDropped(acceptedFiles: any, fileRejections: any, event: any): Promise<void> {
    let sharePointFolderName = "";
    const sharePointStructureEntity = this.cdsService.sharepointSite["folderstructureentity"];
    const primaryEntityFolderUrl =
      this.getSharePointFolderName(
        this.cdsService.primaryEntity[this.cdsService.config.primaryEntityFieldLogicalName],
      ) +
      "_" +
      this.cdsService.config.entityId.replace(/-/g, "").toUpperCase();

    //CASE 1: SharePoint Set to No Primary Entity Structure
    //Example: /sites/Dynamics/lead/Fabrikam_Guid

    if (
      sharePointStructureEntity.toLocaleLowerCase() === "none" ||
      sharePointStructureEntity === undefined ||
      sharePointStructureEntity === "" ||
      sharePointStructureEntity.length === 0
    ) {
      console.log("Using SharePoint Type: 1");
      sharePointFolderName += this.cdsService.config.entityName + "/" + primaryEntityFolderUrl;
      const folderResult = await this.sharePointService.sharepointFolderExists(sharePointFolderName);
      if (!folderResult) {
        await this.sharePointService.createSharePointFolder(this.cdsService.config.entityName, primaryEntityFolderUrl);
      }
    } else {
      //CASE 2: SharePoint Set to Primary Entity Structure & Use Relationship
      //Example /sites/Dynamics/account/Fabrikam_Guid/opportunity/OpportunityName_Guid

      if (
        this.cdsService.config.useRelationship &&
        this.cdsService.parentEntity !== undefined &&
        this.cdsService.parentEntity[sharePointStructureEntity + "id"] !== undefined &&
        this.cdsService.config.relationshipLogicalName
      ) {
        console.log("Using SharePoint Type: 2");
        const parentEntityFolderUrl = this.getSharePointFolderName(
          this.cdsService.parentEntity[this.cdsService.config.relationshipLogicalName] +
            "_" +
            this.cdsService.parentEntity[sharePointStructureEntity + "id"].replace(/-/g, "").toUpperCase(),
        );
        sharePointFolderName +=
          sharePointStructureEntity +
          "/" +
          parentEntityFolderUrl +
          "/" +
          this.cdsService.config.entityName +
          "/" +
          primaryEntityFolderUrl;

        // //Check Folder Tree
        let folderTreeParent = sharePointStructureEntity;
        let folderResult = await this.sharePointService.sharepointFolderExists(
          folderTreeParent + "/" + parentEntityFolderUrl,
        );
        if (!folderResult) {
          console.log("Creating folder: " + folderTreeParent + "/" + parentEntityFolderUrl);
          await this.sharePointService.createSharePointFolder(folderTreeParent, parentEntityFolderUrl);
        }
        folderTreeParent += "/" + parentEntityFolderUrl;
        folderResult = await this.sharePointService.sharepointFolderExists(
          folderTreeParent + "/" + this.cdsService.config.entityName,
        );
        if (!folderResult) {
          console.log("Creating folder: " + folderTreeParent + "/" + this.cdsService.config.entityName);
          await this.sharePointService.createSharePointFolder(folderTreeParent, this.cdsService.config.entityName);
        }
        folderTreeParent += "/" + this.cdsService.config.entityName;
        folderResult = await this.sharePointService.sharepointFolderExists(
          folderTreeParent + "/" + primaryEntityFolderUrl,
        );
        if (!folderResult) {
          console.log("Creating folder: " + folderTreeParent + "/" + primaryEntityFolderUrl);
          await this.sharePointService.createSharePointFolder(folderTreeParent, primaryEntityFolderUrl);
        }
      } else {
        //CASE 3: SharePoint Set to Primary Entity Structure but No Relationship for Upload
        //Example /sites/Dynamics/account/Fabrikam_Guid
        console.log("Using SharePoint Type: 3");
        sharePointFolderName += sharePointStructureEntity + "/" + primaryEntityFolderUrl;
        const folderResult = await this.sharePointService.sharepointFolderExists(sharePointFolderName);
        if (!folderResult) {
          console.log(
            "Creating folder: " +
              this.sharePointRelativeUrl +
              "/" +
              sharePointStructureEntity +
              "/" +
              primaryEntityFolderUrl,
          );
          await this.sharePointService.createSharePointFolder(sharePointStructureEntity, primaryEntityFolderUrl);
        }
      }
    }

    this.droppedFiles = acceptedFiles.length;
    for (let i = 0; i < acceptedFiles.length; i++) {
      const file = acceptedFiles[i] as any;
      debugger;
      if (file.size <= 10485760) {
        await this.sharePointService.web
          .getFolderByServerRelativeUrl(sharePointFolderName)
          .files.add(file.name, file, true);
      } else {
        await this.sharePointService.web.getFolderByServerRelativeUrl(sharePointFolderName).files.addChunked(
          file.name,
          file,
          (data: any) => {
            console.log(data);
          },
          true,
        );
      }
    }
    this.droppedFiles = 0;
  }

  getSharePointFolderName = (input: string): string => {
    const formattedFolderName = input.replace(/[~.{}|&;$%@"?#<>+]/g, "-");
    return formattedFolderName;
  };
}

decorate(SharePointUploadControlVM, {
  isLoading: observable,
  isEnabled: observable,
  droppedFiles: observable,
  currentFile: observable,
  filesProcessed: observable,
  isProcessingFiles: observable,
  onLoad: action.bound,
  onInParametersChanged: action.bound,
  onFileDropped: action.bound,
});
