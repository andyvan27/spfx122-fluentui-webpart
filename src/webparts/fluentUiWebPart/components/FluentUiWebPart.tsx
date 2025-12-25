import * as React from 'react';
import styles from './FluentUiWebPart.module.scss';
import { FluentProvider, teamsLightTheme, teamsDarkTheme, webLightTheme, webDarkTheme } from '@fluentui/react-components';
import type { IFluentUiWebPartProps } from './IFluentUiWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ToyList } from './ToyList';

export default class FluentUiWebPart extends React.Component<IFluentUiWebPartProps> {
  public render(): React.ReactElement<IFluentUiWebPartProps> {
    const {
      listTitle,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      context,
    } = this.props;

    return (
      <FluentProvider theme={hasTeamsContext ? isDarkTheme ? teamsDarkTheme : teamsLightTheme : isDarkTheme ? webDarkTheme : webLightTheme}>
        <section className={`${styles.fluentUiWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
          <ToyList
            listTitle={escape(listTitle)}
            environmentMessage={escape(environmentMessage)}
            userDisplayName={escape(userDisplayName)}
            context={context}
          />
        </section>
      </FluentProvider>
    );
  }
}
