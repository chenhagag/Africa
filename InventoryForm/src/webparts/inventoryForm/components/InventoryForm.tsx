import * as React from 'react';
import styles from './InventoryForm.module.scss';
import type { IInventoryFormProps } from './IInventoryFormProps';

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { MSGraphClientV3 } from '@microsoft/sp-http';



export interface IInventoryItem {
  id: number;
  title: string;
  type: string;
  assignedTo?: string;
  serialNumber?: string;
  available: boolean;
  isNew?: boolean;
  isReturning?: boolean;
  returnReason?: string;
}

export interface IInventoryFormState {
  showForm: boolean;
  selectedUser: { id: number, text: string } | null;
  searchText: string;
  suggestedUsers: any[];
  allUsers: any[];
  allItems: IInventoryItem[];
  userItems: IInventoryItem[];
  isAddingItem: boolean;
  selectedItemType: string;
  availableItems: IInventoryItem[];
  newItems: IInventoryItem[];
  removedItems: IInventoryItem[];
}

export default class InventoryForm extends React.Component<IInventoryFormProps, IInventoryFormState> {
  
  constructor(props: IInventoryFormProps) {
    super(props);   
   
    this.state = {
      showForm: false,
      selectedUser: null,
      allItems: [],
      userItems: [],
      isAddingItem: false,
      selectedItemType: '',
      availableItems: [],
      newItems: [],
      removedItems: [],
      searchText: '',
      suggestedUsers: [],
      allUsers: []
    };
  }


  public componentDidMount(): void {

    this.loadAllUsers();
    this.GetFullInventory();

    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css';
    link.crossOrigin = 'anonymous';
    document.head.appendChild(link);
  }

  private _handleUserInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const value = event.target.value;
    this.setState({ searchText: value });
  
    if (!value) {
      this.setState({ suggestedUsers: [] });
      return;
    }
  
    const filteredUsers = this.state.allUsers.filter(user => {
      const searchLower = value.toLowerCase();
      return (
        (user.displayName && user.displayName.toLowerCase().includes(searchLower)) ||
        (user.mail && user.mail.toLowerCase().includes(searchLower)) ||
        (user.userPrincipalName && user.userPrincipalName.toLowerCase().includes(searchLower))
      );
    });
  
    this.setState({ suggestedUsers: filteredUsers });
  };
  

  private async loadAllUsers(): Promise<void> {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
  
      let allUsers: any[] = [];
      let nextLink: string | undefined = '/users?$select=id,displayName,mail,userPrincipalName,assignedLicenses&$top=999';
      
      while (nextLink) {
        const response: any = await client.api(nextLink).version('v1.0').get(); // <<< הוספתי :any
        allUsers = allUsers.concat(response.value);
        nextLink = response['@odata.nextLink'] ? response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : undefined;
      }
      
  
      const licensedUsers = allUsers.filter((user: any) => user.assignedLicenses && user.assignedLicenses.length > 0);
  
      this.setState({ allUsers: licensedUsers });
  
    } catch (error) {
      console.error('שגיאה בטעינת יוזרים מה-Graph:', error);
    }
  }
  
  



  private async GetFullInventory() {
    try {
      const items = await sp.web.lists.getByTitle("מלאי").items.select(
        "ID",
        "Title",
        "Item/Title","Item/TypeCTT",
        "CurrLand/userCTT",
        "ItemID",
        "Available",
        "CurrOwner/ID","CurrOwner/Title" 
      )
      .expand("CurrLand","Item","CurrOwner").top(4999)
      .get();
  
      const formattedItems: IInventoryItem[] = items.map((item: any) => ({
        id: item.ID,
        title: item.Title,
        type: item["Item"]?.TypeCTT || '',
        assignedTo: item["CurrOwner"]?.Title || '',
        serialNumber: item["ItemID"] || '',
        available: item["Available"] === true,
        isNew: false
      }));

      const availableItems = formattedItems.filter(item => item.available);

      this.setState({
        allItems: formattedItems,
        availableItems: availableItems
      });
      
    } catch (error) {
      console.error("שגיאה בשליפת רשימת מלאי:", error);
    }
  };  


  private ResetForm() {
    this.setState({
      showForm: false,
      allItems: [],
      isAddingItem: false,
      selectedItemType: '',
      availableItems: [],
      newItems: [],
      removedItems: [],
      userItems: []
    });

    this.GetFullInventory();    
  }

  private _handleCancelForm = (): void => {

    this.ResetForm();

    this.setState({
      selectedUser: null,
      searchText: '',
      suggestedUsers: []
    });
  };

  private async saveChanges(): Promise<void> {
    const { userItems, selectedUser } = this.state;
    const today = new Date().toISOString();
  
    for (const item of userItems) {
      if (item.returnReason) {
        // פריט להחזרה
  
        // 1. עדכון ברשימת "מלאי"
        await sp.web.lists.getByTitle("מלאי").items.getById(item.id).update({
          Available: true,
          CurrLandId: null,
          CurrOwnerId: null
        });

        const uniqueTitle = `${selectedUser?.id}-${item.id}`;
  
        // 2. עדכון ברשימת "השאלות"
        const query = await sp.web.lists.getByTitle("השאלות").items
          .filter(`Title eq '${uniqueTitle}'`)
          .top(1)()
          .catch(() => []);
  
        if (query.length > 0) {
          await sp.web.lists.getByTitle("השאלות").items.getById(query[0].ID).update({
            CurrOwnerId: null,
            userCTT: '',
            ReturnDate: today
          });
        }
  
      } else if (item.isNew) {
        const uniqueTitle = `${selectedUser?.id}-${item.id}`;
      
        // 1. עדכון ברשימת "מלאי"
        await sp.web.lists.getByTitle("מלאי").items.getById(item.id).update({
          Available: false,
          CurrOwnerId: selectedUser?.id
        });
      
        // 2. בדיקה אם פריט כבר קיים ברשימת השאלות
        const existing = await sp.web.lists.getByTitle("השאלות").items
          .filter(`Title eq '${uniqueTitle}'`)
          .top(1)()
          .catch(() => []);
      
        if (existing.length === 0) {
          // 3. הוספת פריט חדש
          await sp.web.lists.getByTitle("השאלות").items.add({
            Title: uniqueTitle,
            LandingDate: today,
            CurrOwnerId: selectedUser?.id,
            userCTT: selectedUser?.text,
            ItemId: item.id
          });
        }
      }
      
    }
  
    alert("השינויים נשמרו בהצלחה 🎉");
   
    this.ResetForm(); 
    let currUser = {
      Title: this.state.selectedUser?.text,
      Id: this.state.selectedUser?.id
    }

    
    setTimeout(() => {
      this._handleUserSelect(currUser);
    }, 1000);
  }
  
  
  private _handleAddItemClick = (): void => {
    this.setState({
      isAddingItem: true,
      showForm: true
    });
  };
  

  private _handleItemTypeChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
  
    const selectedType = event.target.value;
  
    if (!selectedType) {
      // אם לא נבחר כלום – נאפס גם את המערך
      this.setState({
        selectedItemType: '',
        availableItems: []
      });
      return;
    }
  
    // סינון מתוך allItems לפי סוג וזמינות
    const filteredItems = this.state.allItems.filter(item =>
      item.type === selectedType && item.available === true
    );
  
    this.setState({
      selectedItemType: selectedType,
      availableItems: filteredItems
    });
  };
    
   
 private addSelectedItem() {

  if (!this.currSelectedItemId) return;
  
  let selectedItem: IInventoryItem | undefined;

  for (let i = 0; i < this.state.availableItems.length; i++) {
    if (this.state.availableItems[i].id === this.currSelectedItemId) {
      selectedItem = this.state.availableItems[i];
      break;
    }
  }
 if (!selectedItem) return;

  // לבדוק אם כבר הוסף לרשימה
  const alreadyExists = this.state.userItems.some(item => item.id === this.currSelectedItemId);
  if (alreadyExists) return;

  const newItem = {
    ...selectedItem,
    isNew: true
  };

  this.setState(prevState => ({
    userItems: [...prevState.userItems, newItem],
    isAddingItem: false,
    selectedItemType: '',
    availableItems: []
  }));

 } 

 private currSelectedItemId = 0;

 private _handleItemSelect = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.currSelectedItemId = parseInt(event.target.value);
  
  };
  
  private _handleRemoveItem(item: IInventoryItem): void {

    const updatedUserItems = this.state.userItems.filter(i => i.id !== item.id);

    let updatedAvailableItems = this.state.availableItems;
  
    if (item.available) {
      updatedAvailableItems = [...this.state.availableItems, { ...item, isNew: false }];
    }
  
    this.setState({
      userItems: updatedUserItems,
      availableItems: updatedAvailableItems
    });
  }
  
  private _handleReturnItem = (item: IInventoryItem): void => {
    const updatedItems = this.state.userItems.map(i =>
      i.id === item.id ? { ...i, isReturning: true } : i
    );
  
    this.setState({ userItems: updatedItems, showForm: true });
  };

  private _handleReturnReasonChange = (itemId: number, reason: string): void => {
    const updatedItems = this.state.userItems.map(i =>
      i.id === itemId ? { ...i, returnReason: reason } : i
    );
  
    this.setState({ userItems: updatedItems });
  };
  

  private _handleSaveReturn = (itemId: number): void => {
    const updatedItems = this.state.userItems.map(i =>
      i.id === itemId
        ? {
            ...i,
            isReturning: false
          }
        : i
    );
  
    this.setState({ userItems: updatedItems });
  };
  
  private _handleCancelReturn = (itemId: number): void => {
    const updatedItems = this.state.userItems.map(i =>
      i.id === itemId ? { ...i, isReturning: false, returnReason: undefined } : i
    );
  
    this.setState({ userItems: updatedItems });
  };
  
  
  private _handleUserSelect = (user: any) => {
    this.setState({
      selectedUser: { id: user.id, text: user.displayName },
      searchText: user.displayName,
      suggestedUsers: []
    });
  
    const filteredItems = this.state.allItems.filter(item => item.assignedTo === user.displayName);
    this.setState({ userItems: filteredItems });
  };
  
  
  public render(): React.ReactElement<IInventoryFormProps> {
     

    return (
      <section className={styles.inventoryForm}>

          <div>
          <h2>טופס ניהול מלאי</h2>

          <div className={styles.formGroup}>
        <label>בחר משתמש:</label>
        <input
          type="text"
          className={styles.userSearch}
          value={this.state.searchText}
          onChange={this._handleUserInputChange}
          placeholder="הקלד שם משתמש"
        />
        {this.state.suggestedUsers.length > 0 && (
          <ul className={styles.userDropdown}>
          {this.state.suggestedUsers.map(user => (
            <li
              key={user.id}
              onClick={() => this._handleUserSelect(user)}
            >
              {user.displayName} ({user.mail})
            </li>
          ))}
          </ul>
        )}
      </div>



          {this.state.selectedUser ? <div>
    {/* טבלת פריטים */}
      <table>
        <thead>
          <tr>
            <th>פריט</th>
            <th>סוג פריט</th>
            <th>פעולות</th>
          </tr>
        </thead>
        <tbody>
          {this.state.userItems.map((item, index) => (
            <tr key={index}>
              <td>{item.title}</td>
              <td>{item.type}</td>
              <td>
                {item.isNew ? (
                  <button
                    className={styles.removeBtn}
                    title="הסר פריט"
                    onClick={() => this._handleRemoveItem(item)}
                  >
                    <i className={`fa-solid fa-trash ${styles.iconTrash}`}></i>
                  </button>
                ) : (

                  <div>
                    {item.isReturning ? (
                        <div className={styles.returnControls}>
                          <select
                            onChange={(e) => this._handleReturnReasonChange(item.id, e.target.value)}
                            value={item.returnReason || ''}
                          >
                            <option value="">בחר סיבה</option>
                            <option value="הפריט הוחזר">הפריט הוחזר</option>
                            <option value="הפריט אבד">הפריט אבד</option>
                            <option value="הפריט נקנה על ידי השואל">הפריט נקנה על ידי השואל</option>
                          </select>
                          <button
                            className={styles.iconButton}
                            onClick={() => this._handleSaveReturn(item.id)}
                            title="שמור"
                          >
                            <i className="fa-solid fa-check"></i>
                          </button>
                          <button
                            className={styles.iconButton}
                            onClick={() => this._handleCancelReturn(item.id)}
                            title="בטל"
                          >
                            <i className="fa-solid fa-xmark"></i>
                          </button>
                        </div>
                      ) : item.returnReason ? (
                        <span className={styles.returnedLabel}>{item.returnReason}</span>
                      ) : (
                        <button
                          className={styles.returnBtn}
                          title="החזר פריט"
                          onClick={() => this._handleReturnItem(item)}
                        >
                          <i className={`fa-solid fa-rotate-left ${styles.iconReturn}`}></i>
                          <span>החזר פריט</span>
                        </button>
                        )}
                  </div>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>

          {/* הוספת פריט */}
          <button className={styles.addItemBtn} onClick={this._handleAddItemClick}>
            + השאלה חדשה
          </button>

          </div>
          : <div/>}
          

          {this.state.isAddingItem ?  <div className={styles.addSection}>
            <div className={styles.formGroup}>
              <label htmlFor="item-type">סוג מוצר:</label>
              <select id="item-type" onChange={this._handleItemTypeChange}>
                <option value="">בחר סוג</option>
                <option>מחשב נייד</option>
                <option>מחשב נייח</option>
                <option>סלולר</option>
              </select>
            </div>

            <div className={styles.formGroup}>
              <label htmlFor="item-select">בחר מוצר:</label>
              <select id="item-select" onChange={this._handleItemSelect}>
                <option value="">בחר מוצר</option>
                {this.state.availableItems.map(item => (
                    <option key={item.id} value={item.id}>{item.title}</option>
                  ))}
              </select>
              <button onClick={() => this.addSelectedItem()} className={styles.addBtn}  >הוסף</button>
            </div>
          </div> : <div/>}
         

          {!this.state.showForm ? (
             <div></div>
          ) : (            
            <div className={styles.formActions}>
              <button className={styles.saveBtn} onClick={() => this.saveChanges()}>
              שמור
            </button>
              <button className={styles.cancelBtn} onClick={this._handleCancelForm}>
                בטל
              </button>
            </div>
            )}

        </div>
       

      </section>
    );
  }
  
}
