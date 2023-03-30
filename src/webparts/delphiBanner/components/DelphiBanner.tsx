import * as React from 'react';
import styles from './DelphiBanner.module.scss';
import { IDelphiBannerProps } from './IDelphiBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class DelphiBanner extends React.Component<IDelphiBannerProps, {}> {
  public render(): React.ReactElement<IDelphiBannerProps> {
    const {
      description,
      bannerImageUrl,
      headerText,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.delphiBanner} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.container}>
          <div>
              <img src={bannerImageUrl} alt='Banner' style={{width:"100%"}} />
          </div>
          <div className={styles.centered}>
              <div>
                <h1>
                    <span style={{fontWeight:"700"}}>{headerText} {escape(userDisplayName)}</span>
                </h1>
                <p>
                  {escape(description)}
                </p>
              </div>
          </div>
        </div>
     </section>
    );
  }
}
