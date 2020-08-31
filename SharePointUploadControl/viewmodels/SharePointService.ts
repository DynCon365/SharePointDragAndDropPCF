/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { sp, containsInvalidFileFolderChars, IFolderAddResult } from "@pnp/sp/presets/all";
import { Web, IWeb } from "@pnp/sp/webs";
import * as Msal from "msal";
import { PnPFetchClient } from "../PnPFetchClient";
import { debug } from "console";

export class SharePointService {
  static serviceProviderName = "SharePointService";
  context: ComponentFramework.Context<unknown>;
  config: {
    sharepointSiteId: string;
    sharePointStructureEntity: string;
    clientId: string;
  };
  msalInstance: Msal.UserAgentApplication;
  web: IWeb;
  sharePointRelativeUrl: string;
  sharePointStructureEntity: string;
  sharePointAboluteUrl: string;
  msalConfig: Msal.Configuration;
  ssoRequest: Msal.AuthenticationParameters;
  constructor(
    context: ComponentFramework.Context<unknown>,
    config: {
      sharepointSiteId: string;
      sharePointStructureEntity: string;
      clientId: string;
    },
  ) {
    this.context = context;
    this.config = config;
  }

  async setupSharePoint(
    loginHint: string,
    sharePointAboluteUrl: string,
    sharePointStructureEntity: string,
  ): Promise<void> {
    const msalConfig = {
      auth: {
        clientId: this.config.clientId,
        authority: "https://login.microsoftonline.com/common",
      },
      cache: {
        storeAuthStateInCookie: true, // Set this to "true" if you are having issues on IE11 or Edge
      },
    };
    const ssoRequest: Msal.AuthenticationParameters = {
      loginHint: loginHint,
    };
    this.msalConfig = msalConfig;
    this.ssoRequest = ssoRequest;
    this.sharePointAboluteUrl = sharePointAboluteUrl;
    const msalInstance = new Msal.UserAgentApplication(msalConfig);
    await msalInstance.ssoSilent(ssoRequest).then((response) => {
      sp.setup({
        sp: {
          // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
          fetchClientFactory: () => {
            return new PnPFetchClient(msalInstance);
          },
        },
      });
    });
    this.web = Web(sharePointAboluteUrl);
    this.sharePointStructureEntity = sharePointStructureEntity;
    this.sharePointRelativeUrl = sharePointAboluteUrl.replace(/^(?:\/\/|[^\/]+)*\//, "");
  }

  async uploadFileToSharePoint(sharePointFolderName: string, file: any): Promise<void> {
    if (file.size <= 10485760) {
      await sp.web.getFolderByServerRelativePath(sharePointFolderName).files.add(file.name, file, true);
    } else {
      await sp.web.getFolderByServerRelativePath(sharePointFolderName).files.addChunked(
        file.name,
        file,
        (data) => {
          console.log(data);
        },
        true,
      );
    }
  }

  async sharepointFolderExists(folderName: string): Promise<boolean> {
    try {
      debugger;
      const folderExists = await this.web.getFolderByServerRelativePath(folderName)();
      return true;
    } catch {
      console.log("SharePoint Folder Not Found - Will be Created");
      return false;
    }
  }

  async createSharePointFolder(folderName: string, newFolder: string): Promise<void> {
    try {
      debugger;
      const folder = await this.web.getFolderByServerRelativePath(folderName).folders.add(newFolder);
    } catch (e) {
      console.log(e);
    }
  }
}
