/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { ServiceProvider } from "./pcf-react/ServiceProvider";
import { SharePointUploadControlVM } from "../viewmodels/SharePointUploadControlVM";
import { ServiceProviderContext } from "../viewmodels/context";
import { Stack, Icon, mergeStyles } from "@fluentui/react";
import Dropzone from "react-dropzone";
import { observer } from "mobx-react";

export interface SharePointUploadControlProps {
  controlWidth?: number;
  controlHeight?: number;
  serviceProvider: ServiceProvider;
}
export interface SharePointControlState {
  hasError: boolean;
}

const iconClass = mergeStyles({
  fontSize: 50,
  height: 50,
  width: 50,
  margin: "0 auto",
});

export class SharePointUploadControl extends React.Component<SharePointUploadControlProps, SharePointControlState> {
  vm: SharePointUploadControlVM;
  constructor(props: SharePointUploadControlProps) {
    super(props);
    this.vm = props.serviceProvider.get<SharePointUploadControlVM>("ViewModel");
    this.state = {
      hasError: false,
    };
  }

  static getDerivedStateFromError(error: any): SharePointControlState {
    // Update state so the next render will show the fallback UI.
    return { hasError: true };
  }

  render(): JSX.Element {
    const { droppedFiles, isLoading, currentFile, isEnabled, onFileDropped, onInParametersChanged } = this.vm;
    return this.state.hasError ? (
      <>Error</>
    ) : (
      <ServiceProviderContext.Provider value={this.props.serviceProvider}>
        <Stack>
          {isLoading === true && (
            <Stack.Item grow>
              <div className={"filesStatsCont uploadDivs"}>
                <div className={"uploadStatusText"}>Loading...</div>
              </div>
            </Stack.Item>
          )}
          <Stack.Item grow>
            <Dropzone onDrop={onFileDropped}>
              {({ getRootProps, getInputProps, isDragActive }): JSX.Element => (
                <section style={{ textAlign: "center" }}>
                  <div {...getRootProps()} style={{ backgroundColor: isDragActive ? "#F8F8F8" : "white" }}>
                    <input {...getInputProps()} />
                    <Icon iconName="CloudUpload" className={iconClass} />
                    <p>Drag &amp; Drop Some Files Here!</p>
                  </div>
                </section>
              )}
            </Dropzone>
          </Stack.Item>
          {droppedFiles > 0 && (
            <Stack.Item grow>
              <div className={"filesStatsCont uploadDivs"} style={{ textAlign: "center" }}>
                <div className={"uploadStatusText"}>
                  Uploading ({currentFile}/{droppedFiles})
                </div>
              </div>
            </Stack.Item>
          )}
        </Stack>
      </ServiceProviderContext.Provider>
    );
  }
}

observer(SharePointUploadControl);
