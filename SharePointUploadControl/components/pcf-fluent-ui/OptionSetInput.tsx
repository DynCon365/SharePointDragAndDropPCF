/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { Dropdown, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import { Stack } from "@fluentui/react";

export interface OptionSetProps {
  value: number | null;
  label?: string;
  options?: ComponentFramework.PropertyHelper.OptionMetadata[];
  showBlank?: boolean;
  blankText?: string;
  onChange: (newValue: number | null) => void;
  dropDownProps?: IDropdownProps;
}

export class OptionSetInput extends React.Component<OptionSetProps> {
  constructor(props: OptionSetProps) {
    super(props);
  }

  onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption | undefined): void => {
    const selectedKey = item?.key == null ? -1 : (item.key as number);
    this.props.onChange(selectedKey != -1 ? selectedKey : null);
  };

  onRenderOption = (option?: IDropdownOption): JSX.Element => {
    const flagIcon = "flag-icon flag-icon-" + option?.text.toLowerCase() ?? "";
    return (
      <div>
        <span className={flagIcon} style={{ paddingLeft: 3, paddingRight: 3 }} />
        <span>{option?.text}</span>
      </div>
    );
  };

  onRenderTitle = (options?: IDropdownOption[]): JSX.Element => {
    let option: IDropdownOption;
    if (options) {
      option = options[0];
      const flagIcon = "flag-icon flag-icon-" + option?.text.toLowerCase() ?? "";
      return (
        <div>
          <span className={flagIcon} style={{ paddingLeft: 5, paddingRight: 5 }} />
          <span>{option?.text}</span>
        </div>
      );
    } else {
      return <div>Error</div>;
    }
  };

  render(): JSX.Element {
    const { options, value, blankText, showBlank, label } = this.props;
    const blankTextLabel = blankText || "--Select--";
    let dropDownOptions =
      (options &&
        options.map((v) => {
          return {
            key: v.Value,
            text: v.Label,
          } as IDropdownOption;
        })) ||
      [];

    if (showBlank) {
      dropDownOptions = [{ key: -1, text: blankTextLabel } as IDropdownOption].concat(dropDownOptions);
    }
    const selectedValue = value == null ? -1 : value;
    return (
      <Stack wrap horizontal tokens={{ maxWidth: 100 }}>
        <Stack.Item>
          <Dropdown
            label={label}
            selectedKey={selectedValue}
            onChange={this.onChange}
            onRenderOption={this.onRenderOption}
            onRenderTitle={this.onRenderTitle}
            options={dropDownOptions || []}
            styles={{ dropdown: { width: 90 } }}
            {...this.props.dropDownProps}
          />
        </Stack.Item>
      </Stack>
    );
  }
}
