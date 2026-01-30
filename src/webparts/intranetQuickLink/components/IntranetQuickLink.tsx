/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/self-closing-comp */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './IntranetQuickLink.module.scss';
import { IIntranetQuickLinkProps, ILinkItem } from './IIntranetQuickLinkProps';
import { SPFx, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Icon } from '@fluentui/react';
import { Placeholder } from '@pnp/spfx-controls-react';

const IntranetQuickLink: React.FC<IIntranetQuickLinkProps> = (props) => {
  const [linkItems, setLinkItems] = useState<ILinkItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const spRef = useRef<any>(null);

  // Default colors from props
  const headerColor = props.headerColor || '#333333';
  const itemBgColor = props.itemBgColor || '#ffffff';
  const itemTextColor = props.itemTextColor || '#333333';
  const itemHoverColor = props.itemHoverColor || '#f3f2f1';
  const iconColor = props.iconColor || '#0078d4';
  const borderColor = props.borderColor || '#e1e1e1';
  const showBorder = props.showBorder !== false; // default to true,
    const headerBgColor = props.headerBgColor || '#f8f9fa';
  const headerTitle = props.headerTitle || 'QUICK LINKS';
  const bodyBgColor = props.bodyBgColor || '#f8f9fa';
  //const bodyTextColor = props.bodyTextColor || '#333333';
  
  // Grid configuration from props
  const maxItems = props.maxItems || 12;
  const itemsPerRow = props.itemsPerRow || 4;

  // Map icon names based on title
  const getIconName = (title: string): string => {
    const titleLower = title.toLowerCase();
    
    const iconMap: { [key: string]: string } = {
      'word': 'WordDocument',
      'excel': 'ExcelDocument',
      'powerpoint': 'PowerPointDocument',
      'teams': 'TeamsLogo',
      'outlook': 'OutlookLogo',
      'onenote': 'OneNoteLogo',
      'sharepoint': 'SharepointLogo',
      'project': 'ProjectLogo',
      'yammer': 'YammerLogo',
      'tutorial': 'ReadingMode',
      'management': 'AccountManagement',
      'add': 'Add',
      'remove': 'Remove',
      'ms': 'MicrosoftLogo',
      'pdf': 'PDF',
      'link': 'Link',
      'document': 'Document',
      'folder': 'FabricFolder',
      'calendar': 'Calendar',
      'mail': 'Mail',
      'people': 'People',
      'task': 'TaskLogo',
      'search': 'Search',
      'settings': 'Settings',
      'help': 'Help'
    };

    for (const key in iconMap) {
      if (titleLower.indexOf(key) !== -1) {
        return iconMap[key];
      }
    }

    return 'Link';
  };

  const loadLinkItems = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const items: ILinkItem[] = await spRef.current.web.lists
        .getByTitle(props.listTitle)
        .items
        .select("Id", "Title", "Link", "Status", "IconName")
        .filter("Status eq 1")
        .orderBy("Created", false)();

      const processedItems = items.map(item => {
        let linkUrl = '';
        
        if (item.Link && typeof item.Link === 'object' && item.Link.Url) {
          linkUrl = item.Link.Url;
        } else if (typeof item.Link === 'string') {
          linkUrl = item.Link;
        }
        
        return {
          ...item,
          Link: linkUrl,
          IconName: item.IconName || getIconName(item.Title || '')
        };
      });
     
      setLinkItems(processedItems);
      setIsLoading(false);
    } catch (error) {
      console.error('Error loading link items:', error);
      setIsLoading(false);
      setErrorMessage(`Failed to load link items.`);
    }
  }, [props.listTitle]);

  useEffect(() => {    
    spRef.current = spfi().using(SPFx(props.context));
    void loadLinkItems();
  }, [props.listTitle, props.context]);

  if (isLoading) {
    return (
      <div className={styles.intranetQuickLink}>
        <div className={styles.container}>
          <div className={styles.loadingContainer}>
            <div className={styles.spinner}></div>
            <div className={styles.loadingText}>Loading quick links...</div>
          </div>
        </div>
      </div>
    );
  }

  if (errorMessage) {
    return (
      <div className={styles.intranetQuickLink}>
        <div className={styles.container}>
          <Placeholder
            iconName='Error'
            iconText='Error'
            description={errorMessage}
          >
            <button
              className={styles.retryButton}
              onClick={() => loadLinkItems()}
            >
              Retry
            </button>
          </Placeholder>
        </div>
      </div>
    );
  }

  // Limit to maxItems
  const displayedItems = linkItems.slice(0, maxItems);

  return (
    
    <div className={styles.intranetQuickLink}
    style={{ backgroundColor: bodyBgColor }} >
      <div className={styles.container}>
        {/* Header */}
        <div className={styles.header}
         style={{ backgroundColor: headerBgColor }}>
          <h1 
            className={styles.title}
            style={{ color: headerColor }}
          >
            {headerTitle}
          </h1>
        </div>

        {/* Grid Container */}
        <div className={styles.gridContainer}>
          {displayedItems.length === 0 ? (
            <div className={styles.noItems}>
              <Icon iconName="Link" className={styles.noItemsIcon} />
              <div className={styles.noItemsText}>No quick links found.</div>
              <div className={styles.noItemsSubtext}>
                Create items in the {props.listTitle} list with Status = Yes.
              </div>
            </div>
          ) : (
            <div 
              className={styles.grid}
              style={{ 
                gridTemplateColumns: `repeat(auto-fill, minmax(${100/itemsPerRow}%, 1fr))`,
                maxHeight: maxItems > 8 ? '500px' : 'auto' // Only add scroll if many items
              }}
            >
              {displayedItems.map((item: ILinkItem) => (
                <a
                  key={item.Id}
                  href={typeof item.Link === 'string' ? item.Link : item.Link?.Url || "#"}
                  className={styles.gridItem}
                  style={{
                    backgroundColor: itemBgColor,
                    color: itemTextColor,
                    '--item-hover-color': itemHoverColor,
                    '--icon-color': iconColor,
                    '--border-color': borderColor,
                    border: showBorder ? `1px solid ${borderColor}` : 'none'
                  } as React.CSSProperties}
                  target='_blank'
                  rel="noopener noreferrer"
                  title={item.Title}
                >
                  <div className={styles.itemIcon}>
                    <Icon 
                      iconName={item.IconName || 'Link'} 
                      className={styles.icon}
                      style={{ color: iconColor }}
                    />
                  </div>
                  <div className={styles.itemTitle}>
                    {item.Title}
                  </div>
                </a>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default IntranetQuickLink;