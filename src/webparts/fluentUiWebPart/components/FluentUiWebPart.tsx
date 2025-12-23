import * as React from 'react';
import styles from './FluentUiWebPart.module.scss';
import { FluentProvider, teamsLightTheme, teamsDarkTheme, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import type { IFluentUiWebPartProps } from './IFluentUiWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ToyList } from './ToyList';

export default class FluentUiWebPart extends React.Component<IFluentUiWebPartProps> {
  public render(): React.ReactElement<IFluentUiWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <FluentProvider theme={hasTeamsContext ? isDarkTheme ? teamsDarkTheme : teamsLightTheme : isDarkTheme ? webDarkTheme : webLightTheme}>
        <section className={`${styles.fluentUiWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
          <ToyList
            description={escape(description)}
            environmentMessage={escape(environmentMessage)}
            userDisplayName={escape(userDisplayName)}
          />
        </section>
      </FluentProvider>
    );
  }
}
