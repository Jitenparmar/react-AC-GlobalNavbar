import * as React from "react";
import styles from '../AppCustomizer.module.scss';
import { ITenantGlobalNavbarProps } from "./ITenantGlobalNavbarProps";
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem, ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';

import * as SPTermStore from './../services/SPTermStoreService'; 

export interface ITenantGlobalNavBarState {
    // So far, it is empty
}

export default class TenantGlobalNavBar extends React.Component<ITenantGlobalNavbarProps,ITenantGlobalNavBarState>{

    /**
 * Main constructor for the component
 */
  constructor(props: ITenantGlobalNavbarProps) {
    super(props);
    this.state = {
    };
  }

  private projectMenuItem(menuItem: SPTermStore.ISPTermObject, itemType: ContextualMenuItemType) : IContextualMenuItem {
    return({
      key: menuItem.identity,
      name: menuItem.name,
      itemType: itemType,
      iconProps:{ iconName: (menuItem.localCustomProperties.iconName != undefined ? menuItem.localCustomProperties.iconName : null)},
      href: menuItem.terms.length == 0 ?
          (menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"] != undefined ?
              menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]
              : null)
          : null,
      subMenuProps: menuItem.terms.length > 0 ? 
          { items : menuItem.terms.map((i) => { return(this.projectMenuItem(i, ContextualMenuItemType.Normal)); }) } 
          : null,
      isSubMenu: itemType != ContextualMenuItemType.Header,
    });
}

  public render(): React.ReactElement<ITenantGlobalNavbarProps>{

    const commandBarItems: IContextualMenuItem[] = this.props.menuItems.map((i) => {
        return(this.projectMenuItem(i, ContextualMenuItemType.Header));
    });

        return(
            <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.app}`}>
                <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.top}`}>
                    <CommandBar
                    className={styles.commandBar}
                    items={ commandBarItems }
                    />
                </div>
            </div>
        );
  }
    
}