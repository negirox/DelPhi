import * as React from 'react';
import styles from './DelPhiQuickLinks.module.scss';
import { IDelPhiQuickLinksProps } from './IDelPhiQuickLinksProps';
import { Icon } from '@fluentui/react/lib/Icon';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../../pnpjsConfig';
import { IDelPhiQuickLinksState } from './IDelPhiQuickLinksState';
import { QLinks } from '../../../models/QLinks';
export default class DelPhiQuickLinks extends React.Component<IDelPhiQuickLinksProps, IDelPhiQuickLinksState> {
  private _sp: SPFI;
  private _tilesImageUrl:string;
  constructor(props: IDelPhiQuickLinksProps, state: IDelPhiQuickLinksState) {
    super(props);
    // set initial state
    this.state = {
      items: Array<QLinks>(),
      errors: []
    };
    this._tilesImageUrl = this.props.tilesImageUrl;
    this._sp = getSP();
    this._readItems();
  }
  private _readItems = async (): Promise<void> => {
    try {
      const allItems = await this._sp.web.lists.getByTitle(this.props.listName).items
                              .select("Title,Id,IconName,QuickLinkUrl,Order,Display")
                              .filter("Display eq 1").orderBy('Order')();
                             console.log(allItems);
       const QItems = allItems.map((item: any) => {
        var qLinks = new QLinks();
        qLinks.IconName = item.IconName;
        qLinks.Id = item.Id;
        qLinks.Title = item.Title;
        qLinks.Link = item.QuickLinkUrl;
        qLinks.Order = item.Order;
        qLinks.Target = item.QuickLinkUrl;
        return qLinks;
      });

      this.setState({ items:QItems});
     
    } catch (err) {
     // Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }
  public render(): React.ReactElement<IDelPhiQuickLinksProps> {
    console.log(this.state.items);
    var imageURl = this._tilesImageUrl;
    return (
      <section className={`${styles.delPhiQuickLinks}}`}>
       <div className={styles.container} >
        <div className={styles.row} >
          {
          this.state.items.map(function (x,k){
            return(
                  <div className={`${styles['col-md-3']}`} >
                    <div className={styles.context}>
                      <img src={imageURl} alt="Icon" className={`${styles['img-fluid']} ${styles['mx-auto']} ${styles['d-block']}`} />
                      <div className={`${styles.overlay} ${styles.green}`}>
                        <div className={styles.text}>
                          <div>
                            <p>
                              <Icon iconName={x.IconName} />
                            </p>
                          </div>
                        </div>
                        </div>
                    </div>
                    <div style={{textAlign:'center'}}>
                      <p>{x.Title}</p>
                    </div>	
                  </div>
              )
            }
          )}
         
        </div>
      </div>
      </section>
    );
  }
}
