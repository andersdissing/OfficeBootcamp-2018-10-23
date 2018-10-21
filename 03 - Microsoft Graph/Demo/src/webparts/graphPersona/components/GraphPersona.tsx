import * as React from 'react';
import styles from './GraphPersona.module.scss';
import { IGraphPersonaProps } from './IGraphPersonaProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IGraphPersonaState } from './IGraphPersonaState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/components/Persona';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
initializeIcons();

export default class GraphPersona extends React.Component<IGraphPersonaProps, IGraphPersonaState> {
  constructor(props: IGraphPersonaProps) {
    super(props);
  
    this.state = {
      name: '',
      email: '',
      phone: '',
      image: null
    };
  }

  public componentDidMount(): void {
    this.props.service.getProfileInfo().then (user =>
    {
        this.setState({
        name: user.displayName,
        email: user.mail,
        phone: user.businessPhones[0]
        });
    });

    this.props.service.getPhoto().then( photoUrl => 
    {
        this.setState({ image: photoUrl });
    });
  }

  private _renderMail = () => {
    if (this.state.email) {
      return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>;
    } else {
      return <div />;
    }
  }
  
  private _renderPhone = () => {
    if (this.state.phone) {
      return <Link href={`tel:${this.state.phone}`}>{this.state.phone}</Link>;
    } else {
      return <div />;
    }
  }

  public render(): React.ReactElement<IGraphPersonaProps> {
    return (
      <div>
        <Persona text={this.state.name}
                secondaryText={this.state.email}
                onRenderSecondaryText={this._renderMail}
                tertiaryText={this.state.phone}
                onRenderTertiaryText={this._renderPhone}
                imageUrl={this.state.image}
                size={PersonaSize.size100} />
        <PrimaryButton onClick={() => this.setState({name: "Per Jakobsen"})} text="Beautify name" />
      </div>
    );
  }
}
