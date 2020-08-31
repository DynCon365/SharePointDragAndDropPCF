/* eslint-disable @typescript-eslint/ban-ts-ignore */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { StandardControlReact } from "./components/pcf-react/StandardControlReact";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { SharePointUploadControl } from "./components/SharePointUploadControl";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { SharePointUploadControlVM } from "./viewmodels/SharePointUploadControlVM";
import { CdsService } from "./viewmodels/CdsService";
import { initializeIcons } from "@uifabric/icons";
import { SharePointService } from "./viewmodels/SharePointService";
initializeIcons();

export class SharePointAttachmentUploadControl extends StandardControlReact<IInputs, IOutputs> {
  constructor() {
    super(true);
    this.renderOnParametersChanged = false;
    this.initServiceProvider = (serviceProvider): void => {
      serviceProvider.register(
        "CdsService",
        new CdsService(this.context, {
          sharePointSiteId: this.context.parameters.sharePointSiteGuid.raw ?? "",
          useRelationship: this.context.parameters.useRelationship.raw === "true",
          //@ts-ignore
          entityId: this.context.mode.contextInfo.entityId,
          //@ts-ignore
          entityName: this.context.mode.contextInfo.entityTypeName,
          primaryEntityFieldLogicalName: this.context.parameters.primaryEntityFieldLogicalName.raw ?? "",
          relationshipLogicalName: this.context.parameters.relationshipLogicalName.raw ?? undefined,
          parentEntityTypeName: this.context.parameters.parentEntityTypeName.raw ?? undefined,
          parentEntityFieldLogicalName: this.context.parameters.parentEntityPrimaryFieldName.raw ?? "",
        }),
      );
      serviceProvider.register(
        "SharePointService",
        new SharePointService(this.context, {
          sharePointStructureEntity: this.context.parameters.parentEntityTypeName.raw ?? "",
          sharepointSiteId: this.context.parameters.sharePointSiteGuid.raw ?? "",
          clientId: this.context.parameters.clientId.raw ?? "",
        }),
      );
      serviceProvider.register("ViewModel", new SharePointUploadControlVM(serviceProvider));
    };
    this.reactCreateElement = (container, width, height, serviceProvider): void => {
      ReactDOM.render(
        React.createElement(SharePointUploadControl, {
          serviceProvider: serviceProvider,
          controlWidth: width,
          controlHeight: height,
        }),
        container,
      );
    };
  }
}
