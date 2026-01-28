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

  // Set default colors if not provided
  const headerColor = props.headerColor || '#333333';
  const rowColor1 = props.rowColor1 || '#ffffff';
  const rowColor2 = props.rowColor2 || '#f8f9fa';
  const rowTextColor = props.rowTextColor || '#333333';
  const rowHoverColor1 = props.rowHoverColor1 || '#f3f2f1';
  const rowHoverColor2 = props.rowHoverColor2 || '#f3f2f1';
  const maxRows = props.maxRows || 4;

  const loadLinkItems = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const items: ILinkItem[] = await spRef.current.web.lists
        .getByTitle(props.listTitle)
        .items
        .select("Id", "Title", "Link", "Status")
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
          Link: linkUrl
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

  // Limit to maxRows
  const displayedItems = linkItems.slice(0, maxRows);

  return (
    <div className={styles.intranetQuickLink}>
      <div className={styles.container}>
        {/* Apply header color */}
        <h1 
          className={styles.title}
          style={{ color: headerColor }}
        >
          QUICK LINKS
        </h1>

        <div className={styles.linksList}>
          {displayedItems.length === 0 ? (
            <div className={styles.noItems}>
              <Icon iconName="Link" className={styles.noItemsIcon} />
              <div className={styles.noItemsText}>No quick links found.</div>
              <div className={styles.noItemsSubtext}>
                Create items in the {props.listTitle} list with Status = Yes.
              </div>
            </div>
          ) : (
            displayedItems.map((item: ILinkItem, index: number) => {
              // Determine if row is even or odd for alternating colors
              const isEvenRow = index % 2 === 0;
              const rowBgColor = isEvenRow ? rowColor1 : rowColor2;
              const rowHoverColor = isEvenRow ? rowHoverColor1 : rowHoverColor2;

              return (
                <a
                  key={item.Id}
                  href={typeof item.Link === 'string' ? item.Link : item.Link?.Url || "#"}
                  className={styles.linkItem}
                  style={{
                    backgroundColor: rowBgColor,
                    color: rowTextColor,
                    '--row-hover-color': rowHoverColor
                  } as React.CSSProperties}
                  target='_blank'
                  rel="noopener noreferrer"
                >
                  {item.Title}
                </a>
              );
            })
          )}
        </div>
      </div>
    </div>
  );
}

export default IntranetQuickLink;