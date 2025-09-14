import * as React from 'react';
//import styles from './SpFxPowerAutomate.module.scss';
import type { ISpFxPowerAutomateProps } from './ISpFxPowerAutomateProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import FlowTrigger from './FlowTrigger';

export default class SpFxPowerAutomate extends React.Component<ISpFxPowerAutomateProps, {}> {
  public render(): React.ReactElement<ISpFxPowerAutomateProps> {
    const {
      context
    } = this.props;

    return (
      <>
      <FlowTrigger context={context} />
      </>
    );
  }
}
