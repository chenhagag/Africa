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
  returnNotes?: string;
  accessories?: { name: string; imageFile?: File; imageUrl?: string; isSelected?: boolean }[];

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
  isSaving: boolean;
  printItems: IInventoryItem[]; 
  printUserName: string;
  subscriptionNumber: string;
  linkedAccessories: { name: string; imageFile?: File; imageUrl?: string; isSelected?: boolean }[];
  isUnlicensed?: boolean;

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
      allUsers: [],
      isSaving: false,
      printItems: [],
      printUserName: '',
      subscriptionNumber: '',
      linkedAccessories: [],
      isUnlicensed: false
    };
  }


  public componentDidMount(): void {

    debugger;
    this.loadAllUsers(false);
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
  

  private async loadAllUsers(showUnlicensed: boolean): Promise<void> {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
  
      let allUsers: any[] = [];
      let nextLink: string | undefined = '/users?$select=id,displayName,mail,userPrincipalName,assignedLicenses&$top=999';
      
      while (nextLink) {
        const response: any = await client.api(nextLink).version('v1.0').get(); // <<< הוספתי :any
        allUsers = allUsers.concat(response.value);
        nextLink = response['@odata.nextLink'] ? response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : undefined;
      }
      
    debugger;
 
      if(showUnlicensed) {
        this.setState({ allUsers: allUsers, isUnlicensed: true });
      } else {    
        const licensedUsers = allUsers.filter((user: any) => user.assignedLicenses && user.assignedLicenses.length > 0);
          this.setState({ allUsers: licensedUsers , isUnlicensed: false });
      }
    
    } catch (error) {
      console.error('שגיאה בטעינת יוזרים מה-Graph:', error);
    }
  }

  private async GetFullInventory() {

    try {
      const items = await sp.web.lists.getByTitle("מלאי").items.select(
        "ID",
        "Title",
        "Item/Title","ItemType",
        "ItemID",
        "Available",
        "CurrOwner/ID","CurrOwner/Title" 
      )
      .expand("Item","CurrOwner").top(4999)
      .get();
  
      const formattedItems: IInventoryItem[] = items.map((item: any) => ({
        id: item.ID,
        title: item.Title,
        type: item.ItemType || '',
        assignedTo: item["CurrOwner"]?.Title || '',
        serialNumber: item["ItemID"] || '',
        available: !item["CurrOwner"]?.Title,
        isNew: false
      }));

      debugger;
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
      userItems: [],
       linkedAccessories: []
    });

    this.GetFullInventory();    
  }

 private _handleCancelForm = (): void => {
  const confirmCancel = window.confirm("האם את בטוחה שברצונך לבטל את הטופס? כל השינויים יאבדו.");

  if (!confirmCancel) {
    return; 
  }

  this.ResetForm();

  this.setState({
    selectedUser: null,
    searchText: '',
    suggestedUsers: []
  });
};

  private async saveChanges(): Promise<void> {


    if (
      this.state.selectedItemType === 'סלולרי' &&
      !this.state.subscriptionNumber.trim()
    ) {
      alert('יש להזין מספר מנוי עבור מכשיר סלולרי');
      return;
    }

    this.setState({ isSaving: true });

    debugger;
    const { userItems, selectedUser } = this.state;
    const today = new Date().toISOString();
  
    for (const item of userItems) {

      debugger;
      if (item.returnReason) {
        // פריט להחזרה
        var itemStatus = "נגרע";
        debugger;
        if (item.returnReason == "הפריט הוחזר") 
            itemStatus = 'פנוי להשאלה';

        // 1. עדכון ברשימת "מלאי"
        await sp.web.lists.getByTitle("מלאי").items.getById(item.id).update({
          Available: true,
          CurrLandId: null,
          CurrOwnerId: null,
          ItemStatus: itemStatus,
          ReturnReason: item.returnReason       
         });


        debugger;
  
        // 2. עדכון ברשימת "השאלות"
        const query = await sp.web.lists.getByTitle("השאלות").items
          .filter(`ItemId eq ${item.id} and CurrOwnerId eq ${selectedUser?.id} and Status eq 'השאלה פעילה'`)

          .top(1)()
          .catch(() => []);
  
        if (query.length > 0) {
          await sp.web.lists.getByTitle("השאלות").items.getById(query[0].ID).update({
            Status: 'הוחזר',
            ReturnDate: today,
            ReturnReason: item.returnReason,
            ReturnNotes: item.returnNotes || ''
          });
        }
  
      } else if (item.isNew) {
        const uniqueTitle = `${selectedUser?.text}-${item.id}`;
      
        // 1. עדכון ברשימת "מלאי"
        await sp.web.lists.getByTitle("מלאי").items.getById(item.id).update({
          Available: false,
          CurrOwnerId: selectedUser?.id,
          ItemStatus: 'מושאל',
          ReturnReason: 'הפריט בהשאלה'
        });
  
        // 2. בדיקה אם פריט כבר קיים ברשימת השאלות
        const existing = await sp.web.lists.getByTitle("השאלות").items
  .filter(`Title eq '${uniqueTitle}' and Status eq 'השאלה פעילה'`)
          .top(1)()
          .catch(() => []);
      
if (existing.length === 0) {
  // איסוף מוצרים נלווים מהמוצר עצמו
  const itemAccessories = (item as any).accessories || [];
  const selectedAccessories = itemAccessories
  .filter((acc: { name?: string; isSelected?: boolean }) => acc.isSelected && acc.name?.trim())
  .map((acc: { name: string }) => ({ ...acc, name: acc.name.trim() }));

  const accessoriesText = selectedAccessories.map((acc: { name: string }) => acc.name).join(', ');

  // יצירת פריט חדש ברשימת השאלות
  const addedItem = await sp.web.lists.getByTitle("השאלות").items.add({
    Title: uniqueTitle,
    LandingDate: today,
    CurrOwnerId: selectedUser?.id,
    ItemId: item.id,
    Status: 'השאלה פעילה',
    Accessories: accessoriesText,
    subscription: this.state.subscriptionNumber || null

  });

  const itemId = addedItem.data.ID;
        // הוספת התמונות כקבצים מצורפים
        for (const acc of selectedAccessories) {
          if (acc.imageFile) {
            const safeName = acc.name || `accessory-${Date.now()}`;
            try {
              await sp.web.lists
                .getByTitle("השאלות")
                .items.getById(itemId)
                .attachmentFiles.add(`${safeName}.jpg`, acc.imageFile);
              } catch (err) {
               console.error(`שגיאה בהעלאת קובץ עבור ${safeName}`, err);
              }
            }
          }
        } 
      }
    }
        
    await new Promise(resolve => setTimeout(resolve, 500));

    let currUser = {
      Title: this.state.selectedUser?.text,
      Id: this.state.selectedUser?.id
    };

    setTimeout(() => {
      this.ResetForm(); 
      this.setState({ isSaving: false });
      alert("השינויים נשמרו בהצלחה 🎉");

      setTimeout(() => {
        const filteredItems = this.state.allItems.filter(item => item.assignedTo === currUser.Title);
        this.setState({ 
          userItems: filteredItems,
          printItems: filteredItems,
          printUserName: currUser.Title  || 'משתמש נוכחי',
          subscriptionNumber: ''
        });         
       }, 1000);    
   }, 1000);

  }

  
  private _handleAddItemClick = (): void => {
   
    this.setState({
      isAddingItem: true,
      showForm: true
    });   
  };

  private _handleAddAccessory = (): void => {
  this.setState(prev => ({
    linkedAccessories: [...prev.linkedAccessories, { name: '', isSelected: true }]
  }));
};

  
  private _handlePrint = (): void => {
  const printWindow = window.open('', '_blank');

  if (!printWindow) return;

  const { printUserName, userItems } = this.state;

  const htmlContent = `
    <html dir="rtl" lang="he">
      <head>
        <title>רשימת פריטים</title>
        <style>
          body { font-family: Arial; padding: 20px; }
          h2 { text-align: center; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #999; padding: 8px; text-align: right; }
          th { background-color: #f0f0f0; }
        </style>
      </head>
      <body>
        <h2>פריטים של ${printUserName}</h2>
        <table>
          <thead>
            <tr>
              <th>שם פריט</th>
              <th>סוג</th>
              <th>מספר סידורי</th>
            </tr>
          </thead>
          <tbody>
            ${userItems.map(item => `
              <tr>
                <td>${item.title}</td>
                <td>${item.type}</td>
                <td>${item.serialNumber || ''}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </body>
    </html>
  `;

  printWindow.document.write(htmlContent);
  printWindow.document.close();

  printWindow.onload = () => {
    printWindow.focus();
    printWindow.print();
  };
};


 private _handleItemTypeChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
 
  debugger;
  const selectedType = event.target.value;

  let linkedAccessories: any = [];

  if (selectedType === "מחשב נייד") {
    linkedAccessories = [
      { name: "עכבר" },
      { name: "מקלדת" },
      { name: "מסך" },
      { name: "תחנת עגינה" },
      { name: "רמקול" },
      { name: "מצלמה" }
    ];
  }

  const filteredItems = this.state.allItems.filter(item =>
    item.type === selectedType && item.available === true
  );

  this.setState({
    selectedItemType: selectedType,
    availableItems: filteredItems,
    linkedAccessories: linkedAccessories
  });
};

   
private _handleAccessoryImageChange = (index: number, file: File): void => {
  const reader = new FileReader();
  reader.onloadend = () => {
    const newAccessories = [...this.state.linkedAccessories];
    newAccessories[index].imageFile = file;
    newAccessories[index].imageUrl = reader.result as string;
    this.setState({ linkedAccessories: newAccessories });
  };
  reader.readAsDataURL(file);
};

   
private addSelectedItem() {
  if (!this.currSelectedItemId) return;

  const selectedItem = this.state.availableItems.find(
    item => item.id === this.currSelectedItemId
  );

  if (!selectedItem) return;

  const alreadyExists = this.state.userItems.some(
    item => item.id === this.currSelectedItemId
  );
  if (alreadyExists) return;

  // צרף את המוצרים הנלווים למוצר הזה
  const newItem = {
    ...selectedItem,
    isNew: true,
    accessories: this.state.linkedAccessories // ← כאן השיוך
  };

  const updatedAll = this.state.allItems.filter(
    i => i.id !== this.currSelectedItemId
  );

  this.setState(prevState => ({
    userItems: [...prevState.userItems, newItem],
    isAddingItem: false,
    selectedItemType: '',
    allItems: updatedAll,
    linkedAccessories: [] // ← אפס את הטופס של המוצרים הנלווים
  }));
}


 private currSelectedItemId = 0;

 private _handleItemSelect = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.currSelectedItemId = parseInt(event.target.value);
  
  };
  
private _handleRemoveItem(item: IInventoryItem): void {
  const updatedUserItems = this.state.userItems.filter(i => i.id !== item.id);

  // סינון כפילויות - רק אם עוד לא קיים ב־availableItems
  const alreadyAvailable = this.state.allItems.some(i => i.id === item.id);
  const updatedAvailableItems = alreadyAvailable
    ? this.state.allItems
    : [...this.state.allItems, { ...item, isNew: false }];

  this.setState({
    userItems: updatedUserItems,
    allItems: updatedAvailableItems
  });
}
  
  private _handleReturnItem = (item: IInventoryItem): void => {
  
    const updatedItems = this.state.userItems.map(i =>
    i.id === item.id ? { ...i, isReturning: true } : i
  );

  const updatedAvailableItems = this.state.allItems.filter(i => i.id !== item.id);

  this.setState({
    userItems: updatedItems,
    allItems: updatedAvailableItems,
    showForm: true,
    isAddingItem: false
  });
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
  
  
  private _handleUserSelect = async (user: any) => {
    try {
      const ensuredUser = await sp.web.ensureUser(user.userPrincipalName); 
      const spUserId = ensuredUser.data.Id; 
  
      this.setState({
        selectedUser: { id: spUserId, text: user.displayName },
        searchText: user.displayName,
        suggestedUsers: [],
        printUserName: ''
      });

      const filteredItems = this.state.allItems.filter(item => item.assignedTo === user.displayName);
      this.setState({ userItems: filteredItems });


    } catch (error) {
      console.error('שגיאה בהבאת מזהה SharePoint:', error);
    }

  };


  private _handleReturnNotesChange = (itemId: number, notes: string): void => {
  const updatedItems = this.state.userItems.map(i =>
    i.id === itemId ? { ...i, returnNotes: notes } : i
  );
  this.setState({ userItems: updatedItems });
};

    
  public render(): React.ReactElement<IInventoryFormProps> {
     
    return (
      <section className={styles.inventoryForm}>

          <div>
          <h2>טופס ניהול מלאי</h2>

          {this.state.isSaving && (
            <div className={styles.spinnerOverlay}>
              <div className={styles.spinner}></div>
              <div className={styles.spinnerText}>שומר שינויים...</div>
            </div>
          )}


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

      {this.state.isUnlicensed ? (      <button className={styles.licensedBtn} onClick={() => this.loadAllUsers(false)}>הצג משתמשים פעילים בלבד</button>
      ) : (      <button className={styles.licensedBtn} onClick={() => this.loadAllUsers(true)}>הצג משתמשים שאינם פעילים</button>
      )}      
      



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

                            <textarea
                              placeholder="הקלד הערות..."
                              value={item.returnNotes || ''}
                              onChange={(e) => this._handleReturnNotesChange(item.id, e.target.value)}
                              className={styles.notesTextarea}
                            />

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

            {this.state.printItems.length >= 0 && this.state.printUserName != '' && (
                <button className={styles.printBtn} onClick={this._handlePrint}>
                  הדפס פריטים
                </button>
              )}

          </div>
          : <div/>}
          

          {this.state.isAddingItem ?  <div className={styles.addSection}>


            <div className={styles.formGroup}>
              <label htmlFor="item-type">סוג מוצר:</label>
              <select id="item-type" onChange={this._handleItemTypeChange}>
                <option value="">בחר סוג</option>
                <option>מחשב נייד</option>
                <option>מחשב נייח</option>
                <option>סלולרי</option>
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

              {this.state.selectedItemType === 'סלולרי' && (
              <div className={styles.formGroup}>
                <label htmlFor="subscription-number">מספר מנוי:</label>
                <input
                  type="text"
                  id="subscription-number"
                  value={this.state.subscriptionNumber}
                  onChange={(e) => this.setState({ subscriptionNumber: e.target.value })}
                  placeholder="הכנס מספר מנוי"
                />
              </div>
            )}

              
            </div>
      {(this.state.linkedAccessories.length > 0 ) && (
 
 <div className={styles.accessoriesSection}>
    <h4>מוצרים נלווים</h4>
    {this.state.linkedAccessories.map((acc, index) => (
      <div key={index} className={styles.accessoryRow}>
        <input
          type="checkbox"
          checked={!!acc.isSelected}
          onChange={(e) => {
            const updated = [...this.state.linkedAccessories];
            updated[index].isSelected = e.target.checked;
            this.setState({ linkedAccessories: updated });
          }}
        />
        <input
          type="text"
          value={acc.name}
          onChange={(e) => {
            const updated = [...this.state.linkedAccessories];
            updated[index].name = e.target.value;
            this.setState({ linkedAccessories: updated });
          }}
          placeholder="שם מוצר נלווה"
        />
        <input
          type="file"
          accept="image/*"
          onChange={(e) => {
            if (e.target.files?.length) {
              this._handleAccessoryImageChange(index, e.target.files[0]);
            }
          }}
        />
        {acc.imageUrl && (
          <img src={acc.imageUrl} alt={acc.name} style={{ maxWidth: 80 }} />
        )}
      </div>
    ))}

    <button type="button" onClick={this._handleAddAccessory}>
      + הוסף מוצר נלווה נוסף
    </button>
  </div>
)}

           <button onClick={() => this.addSelectedItem()} className={styles.addBtn}>הוסף</button>           
          </div> : <div/>}



          {!this.state.showForm ? (
             <div></div>
          ) : (            
            <div className={styles.formActions}>
              <button className={styles.saveBtn} onClick={() => this.saveChanges()}
                disabled={this.state.isAddingItem}>
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
