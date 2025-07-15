import * as React from 'react';
import styles from './InvetoryDisplay.module.scss';
import type { IInvetoryDisplayProps } from './IInvetoryDisplayProps';

import { sp } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IInventoryItem {
  id: number;
  title: string;
  type: string;
  assignedTo?: string;
  serialNumber?: string;
  available: boolean;
}

export interface ILoanItem {
  id: number;
  title: string;
  userCTT?: string;
  landingDate?: string;
  returnDate?: string;
  status?: string;
  itemId?: number;
  err?: string; // ← חדש
}

interface IState {
  inventoryItems: IInventoryItem[];
  loanItems: ILoanItem[];
  searchText: string;
  filterType: string;
}

export default class InvetoryDisplay extends React.Component<IInvetoryDisplayProps, IState> {
  constructor(props: IInvetoryDisplayProps) {
    super(props);

    sp.setup({
      spfxContext: this.props.context as any
    });

    this.state = {
      inventoryItems: [],
      loanItems: [],
      searchText: '',
      filterType: ''
    };
  }

  public componentDidMount(): void {
    this.loadInventory();
    this.loadLoans();
  }

  private async loadInventory(): Promise<void> {
    try {
      const items = await sp.web.lists.getByTitle("מלאי").items
        .select("ID", "Title", "ItemType", "ItemID", "CurrOwner/Title")
        .expand("CurrOwner", "Item")
        .top(4999)
        .get();

      const formatted: IInventoryItem[] = items.map((item: any) => ({
        id: item.ID,
        title: item.Title,
        type: item.ItemType || '',
        serialNumber: item.ItemID || '',
        assignedTo: item.CurrOwner?.Title || '',
        available: !item.CurrOwner?.Title
      }));

      this.setState({ inventoryItems: formatted });
    } catch (error) {
      console.error('שגיאה בטעינת מלאי:', error);
    }
  }

  private async loadLoans(): Promise<void> {
    try {
      const items = await sp.web.lists.getByTitle("השאלות").items
        .select("ID", "Title", "userCTT", "LandingDate", "ReturnDate", "Status", "ItemId", "Err")
        .top(4999)
        .get();

      const formatted: ILoanItem[] = items.map((item: any) => ({
        id: item.ID,
        title: item.Title,
        userCTT: item.userCTT,
        landingDate: item.LandingDate,
        returnDate: item.ReturnDate,
        status: item.Status,
        itemId: item.ItemId,
        err: item.Err // ← חדש
      }));

      this.setState({ loanItems: formatted });
    } catch (error) {
      console.error('שגיאה בטעינת השאלות:', error);
    }
  }

  private handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ searchText: event.target.value });
  };

  private applyFilter = (type: string) => {
    this.setState(prev => ({
      filterType: prev.filterType === type ? '' : type
    }));
  };

  public render(): React.ReactElement<IInvetoryDisplayProps> {
    const { inventoryItems, loanItems, searchText, filterType } = this.state;

    const filteredInventory = inventoryItems.filter(item => {
      const matchesSearch = item.title.includes(searchText);

      switch (filterType) {
        case 'loaned':
          return !item.available && matchesSearch;
        case 'available':
          return item.available && matchesSearch;
        case 'counted':
          return item.serialNumber && matchesSearch;
        case 'notCounted':
          return !item.serialNumber && matchesSearch;
        default:
          return matchesSearch;
      }
    });

    const filteredLoans = loanItems.filter(item =>
      item.title.includes(searchText)
    );

    return (
      <div className={styles.invetoryDisplay}>
        <div className={styles.controls}>
          <input
            type="text"
            placeholder="חיפוש..."
            value={searchText}
            onChange={this.handleSearchChange}
          />
          <button onClick={() => this.applyFilter('loaned')}>פריטים מושאלים</button>
          <button onClick={() => this.applyFilter('available')}>פריטים פנויים</button>
          <button onClick={() => this.applyFilter('counted')}>נספר במלאי</button>
          <button onClick={() => this.applyFilter('notCounted')}>לא נספר במלאי</button>
        </div>

        <h3>טבלת מלאי</h3>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>פריט</th>
              <th>סוג</th>
              <th>מס׳ סידורי</th>
              <th>משויך ל</th>
            </tr>
          </thead>
          <tbody>
            {filteredInventory.map(item => (
              <tr key={item.id}>
                <td>{item.title}</td>
                <td>{item.type}</td>
                <td>{item.serialNumber}</td>
                <td>{item.assignedTo || 'פנוי'}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <h3>טבלת השאלות</h3>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>כותרת</th>
              <th>משתמש</th>
              <th>תאריך השאלה</th>
              <th>תאריך החזרה</th>
              <th>סטטוס</th>
            </tr>
          </thead>
          <tbody>
            {filteredLoans.map(item => {
              const hasError = item.err === 'ERROR';
              return (
                <tr
                  key={item.id}
                  style={hasError ? { color: 'red', fontWeight: 'bold' } : {}}
                >
                  <td>
                    {item.title}
                    {hasError && (
                      <span style={{ marginRight: '10px' }}>
                        ⚠ שגיאה בפריט! הפריט שונה לאחר שהוחזר
                      </span>
                    )}
                  </td>
                  <td>{item.userCTT}</td>
                  <td>{item.landingDate}</td>
                  <td>{item.returnDate}</td>
                  <td>{item.status}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  }
}
