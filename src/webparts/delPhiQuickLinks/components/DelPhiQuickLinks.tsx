import * as React from 'react';
import styles from './DelPhiQuickLinks.module.scss';
import { IDelPhiQuickLinksProps } from './IDelPhiQuickLinksProps';
import { Icon } from '@fluentui/react/lib/Icon';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../../pnpjsConfig';
import { IDelPhiQuickLinksState } from './IDelPhiQuickLinksState';
import { QLinks } from '../../../models/QLinks';
import { HelperUtils } from '../../../utils/HelperUtils';
/* import { DefaultPalette, Stack, IStackStyles, IStackTokens } from '@fluentui/react';

// Non-mutating styles definition
const itemStyles: React.CSSProperties = {
  alignItems: 'center',
  background: DefaultPalette.themePrimary,
  color: DefaultPalette.white,
  display: 'flex',
  height: 50,
  justifyContent: 'center',
  width: 50,
};

// Tokens definition
const sectionStackTokens: IStackTokens = { childrenGap: 30 };
const wrapStackTokens: IStackTokens = { childrenGap: 150 };
const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themeTertiary,
    width: `100%`,
  },
}; */

export default class DelPhiQuickLinks extends React.Component<IDelPhiQuickLinksProps, IDelPhiQuickLinksState> {
  private _sp: SPFI;
  private _isBackGround:boolean;
  constructor(props: IDelPhiQuickLinksProps, state: IDelPhiQuickLinksState) {
    super(props);
    // set initial state
    this.state = {
      items: Array<QLinks>(),
      errors: []
    };
    //this._tilesImageUrl = this.props.tilesImageUrl;
    this._isBackGround = this.props.isBackGround;
    //this._backgroundColor = this.props.backGroundColor;
    this._sp = getSP();
    this.onInit();
  }
  private GetBox(obj:QLinks):JSX.Element{
    if(HelperUtils.isEmpty(obj.BackgroundImageUrl) && this._isBackGround){
      return (<div style={{backgroundColor:obj.BackgroundColor, minHeight:'100px'}} className={`${styles['d-block']}`}></div>)
    }
    else{
      return  (<img src={obj.BackgroundImageUrl} alt="Icon" className={`${styles['img-fluid']} ${styles['mx-auto']} ${styles['d-block']}`} />)
    }
  } 
  public async onInit(): Promise<void> {
    console.log('In onInit');
    await this._readItems();
  }
  private _readItems = async (): Promise<void> => {
    try {
      const allItems = await this._sp.web.lists.getByTitle(this.props.listName).items
                              .select("Title,Id,IconName,QuickLinkUrl,Order,Display,BackgroundImageUrl,BackgroundColor")
                              .filter("Display eq 1").orderBy('Order')();
                             console.log(allItems);
       const QItems = allItems.map((item: any) => {
        var qLinks = new QLinks();
        qLinks.IconName = item.IconName === (undefined || null) ? 'link' : item.IconName;
        qLinks.Id = item.Id;
        qLinks.Title = item.Title;
        qLinks.Link = item.QuickLinkUrl;
        qLinks.Order = item.Order;
        qLinks.Target = item.QuickLinkUrl;
        qLinks.BackgroundImageUrl = item.BackgroundImageUrl;
        qLinks.BackgroundColor = item.BackgroundColor;
        return qLinks;
      });

      this.setState({ items:QItems});
     
    } catch (err) {
     // Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }
  public render(): React.ReactElement<IDelPhiQuickLinksProps> {
    console.log(this.state.items);
    const self = this;
    return (
      <section className={`${styles.delPhiQuickLinks}}`}>
       <div className={styles.container} >
        <div className={styles.row} >
          {
          this.state.items.map(function (x,k){
            return(
                  <div className={`${styles['col-md-3']}`} >
                    <div className={styles.context}>
                      {self.GetBox(x)}
                      <div className={`${styles.overlay}`}>
                        <div className={styles.text}>
                          <div>
                            <p>
                             <a href={x.Link} target='_blank' style={{color:'white'}}><Icon iconName={x.IconName} /></a> 
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div style={{textAlign:'center',fontSize:'12px'}}>
                      <p><strong><a href={x.Link} target='_blank' style={{color:'black', textDecoration:'none'}}>{x.Title}</a></strong></p>
                    </div>
                  </div>
              )
            }
          )}
         
        </div>
        {/* <div className={styles.row} >
          {
          this.state.items.map(function (x,k){
            return(
              <Stack enableScopedSelectors tokens={sectionStackTokens}>
                <Stack enableScopedSelectors horizontal wrap styles={stackStyles} tokens={wrapStackTokens}>
                  <p>
                    <span style={itemStyles}> <a href={x.Link} target='_blank'><Icon iconName={x.IconName} /></a> </span>
                    <a href={x.Link} target='_blank'>{x.Title}</a>
                  </p>
                </Stack>
            </Stack>
              )
            }
          )}
         
        </div> */}
      </div>
      </section>
    );
  }
}
