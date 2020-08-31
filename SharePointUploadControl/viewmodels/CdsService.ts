export class CdsService {
  static serviceProviderName = "CdsService";
  context: ComponentFramework.Context<unknown>;
  primaryEntity: ComponentFramework.WebApi.Entity;
  parentEntity: ComponentFramework.WebApi.Entity;
  sharepointSite: ComponentFramework.WebApi.Entity;
  currentUser: ComponentFramework.WebApi.Entity;
  config: {
    entityName: string;
    entityId: string;
    useRelationship: boolean;
    primaryEntityFieldLogicalName: string;
    parentEntityTypeName: string | undefined;
    parentEntityFieldLogicalName: string;
    relationshipLogicalName: string | undefined;
    sharePointSiteId: string;
  };
  constructor(
    context: ComponentFramework.Context<unknown>,
    config: {
      entityName: string;
      entityId: string;
      parentEntityTypeName: string | undefined;
      parentEntityFieldLogicalName: string;
      primaryEntityFieldLogicalName: string;
      useRelationship: boolean;
      sharePointSiteId: string;
      relationshipLogicalName: string | undefined;
    },
  ) {
    this.context = context;
    this.config = config;
  }

  async getPrimaryEntity(): Promise<void> {
    try {
      console.log("Getting Primary Entity");
      this.primaryEntity = await this.context.webAPI.retrieveRecord(this.config.entityName, this.config.entityId);
    } catch (error) {
      const errOptions = { message: error.message };
      this.context.navigation.openErrorDialog(errOptions);
    }
  }

  async getParentEntity(): Promise<void> {
    try {
      if (
        this.config.useRelationship &&
        this.config.parentEntityTypeName &&
        this.primaryEntity !== undefined &&
        this.primaryEntity["_" + this.config.relationshipLogicalName + "_value"] !== undefined &&
        this.primaryEntity["_" + this.config.relationshipLogicalName + "_value"] !== ""
      ) {
        console.log("Getting Parent Entity");
        const relationshipId = this.primaryEntity["_" + this.config.relationshipLogicalName + "_value"];
        this.parentEntity = await this.context.webAPI.retrieveRecord(this.config.parentEntityTypeName, relationshipId);
      } else {
        return undefined;
      }
    } catch (error) {
      const errOptions = { message: error.message };
      this.context.navigation.openErrorDialog(errOptions);
    }
  }

  async getSharePointSite(): Promise<void> {
    try {
      console.log("Getting SharePoint Site");
      this.sharepointSite = await this.context.webAPI.retrieveRecord("sharepointsite", this.config.sharePointSiteId);
    } catch (error) {
      const errOptions = { message: error.message };
      this.context.navigation.openErrorDialog(errOptions);
    }
  }

  async getCurrentUser(): Promise<void> {
    try {
      console.log("Getting User");
      if (this.context.userSettings.userName.length) {
        this.currentUser = await this.context.webAPI.retrieveRecord("systemuser", this.context.userSettings.userId);
      } else {
        const errOptions = { message: "User not found in Context." };
        this.context.navigation.openErrorDialog(errOptions);
      }
    } catch (error) {
      const errOptions = { message: error.message };
      this.context.navigation.openErrorDialog(errOptions);
    }
  }
}
