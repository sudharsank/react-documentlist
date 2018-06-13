import * as React from 'react';
import { IConfigContainerProps } from './IConfigContainerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';

export default class ConfigContainer extends React.Component<IConfigContainerProps, {}> {
  public render(): React.ReactElement<IConfigContainerProps> {
    return (
      <div>
        {this.props.displayButton && 
          <Placeholder
            iconName='Edit'
            iconText={this.props.iconText}
            description={this.props.description}
            buttonLabel={this.props.buttonText}
            onConfigure={this.onConfigure} />
        }
        {!this.props.displayButton && 
          <Placeholder
            iconName='Edit'
            iconText={this.props.iconText}
            description={this.props.description} />
        }
      </div>
    );
  }

  private onConfigure = () => {
    this.props.currentContext.propertyPane.open();
  }
}
