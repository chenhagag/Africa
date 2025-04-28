import * as React from 'react';
import styles from './AfricaDeliveryNewForm.module.scss';

import { sp } from "@pnp/sp/presets/all";
import { IAfricaDeliveryNewFormProps } from './IAfricaDeliveryNewFormProps';

import { MSGraphClientV3 } from '@microsoft/sp-http';


interface IFormData {
  deliveryType: string;
  deliveryTo: string;
  externalDestination: string;
  packageContent: string;
  notes: string;
}

interface IState {
  showForm: boolean;
  formData: IFormData;

  selectedUser: { id: number, text: string } | null;
  searchText: string;
  suggestedUsers: any[];
  allUsers: any[];

}

export default class AfricaDeliveryNewForm extends React.Component<IAfricaDeliveryNewFormProps, IState> {
  constructor(props: IAfricaDeliveryNewFormProps) {
    super(props);
    
    this.state = {
      showForm: false,
      searchText: "",
      selectedUser: null,
      allUsers: [],
      suggestedUsers: [],
      formData: {
        deliveryType: '',
        deliveryTo: '',
        externalDestination: '',
        packageContent: '',
        notes: ''
      }
    };

    sp.setup({
      spfxContext: this.props.context as any
    });
  }


  public componentDidMount(): void {

    this.loadAllUsers();
 }

  private async loadAllUsers(): Promise<void> {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');

      let allUsers: any[] = [];
      let nextLink: string | undefined = '/users?$select=id,displayName,mail,userPrincipalName,assignedLicenses&$top=999';
      
      while (nextLink) {
        const response: any = await client.api(nextLink).version('v1.0').get(); 
        allUsers = allUsers.concat(response.value);
        nextLink = response['@odata.nextLink'] ? response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : undefined;
      }
      

      const licensedUsers = allUsers.filter((user: any) => user.assignedLicenses && user.assignedLicenses.length > 0);

      this.setState({ allUsers: licensedUsers });

    } catch (error) {
      console.error('שגיאה בטעינת יוזרים מה-Graph:', error);
    }
  }

  private _handleUserSelect = async (user: any) => {
    try {
      const ensuredUser = await sp.web.ensureUser(user.userPrincipalName); 
      const spUserId = ensuredUser.data.Id; 
  
      this.setState({
        selectedUser: { id: spUserId, text: user.displayName },
        searchText: user.displayName,
        suggestedUsers: []
      });
    } catch (error) {
      console.error('שגיאה בהבאת מזהה SharePoint:', error);
    }
  };
  

  private openForm = (): void => {
    this.setState({ showForm: true });
  };

  private closeForm = (): void => {
    this.setState({ showForm: false });
  };

  private handleInputChange = (event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
    const { name, value } = event.target;
    this.setState(prevState => ({
      formData: {
        ...prevState.formData,
        [name]: value
      }
    }));
  };

  private handleDeliveryTypeChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const { value } = event.target;
    this.setState(prevState => ({
      formData: {
        ...prevState.formData,
        deliveryType: value
      }
    }));
  };

  private handleSave = async (): Promise<void> => {
    try {
      const { formData, selectedUser } = this.state;
  
      // ולידציה לפני שמירה
      if (!formData.deliveryType) {
        alert('אנא בחר סוג משלוח');
        return;
      }
  
      if (!selectedUser && !formData.externalDestination) {
        alert('אנא בחר נמען למשלוח או הזן יעד חיצוני');
        return;
      }
  
      // שלב ראשון: יצירת פריט חדש
      const addResult = await sp.web.lists.getByTitle("שליחויות").items.add({
        Title: "משלוח בהנפקה", 
        Importance: formData.deliveryType,
        DeliverToId: selectedUser?.id,
        DestinationTxt: formData.externalDestination,
        PackageContent: formData.packageContent,
        Notes: formData.notes
      });
  
      // שלב שני: עדכון מספר משלוח עם ה-ID
      const newItemId = addResult.data.Id;
      const shippingNumber = `111${newItemId}`;
  
      // בחירת טקסט עבור TitleForApp
      const forWhomText = selectedUser?.text 
        ? selectedUser.text 
        : formData.externalDestination;
  
      await sp.web.lists.getByTitle("שליחויות").items.getById(newItemId).update({
        barcode: shippingNumber,
        Title: "משלוח מספר " + shippingNumber, 
        TitleForApp: "משלוח מספר " + shippingNumber + " עבור " + forWhomText,
      });
  
      alert("נשמר בהצלחה!");
      window.location.reload(); 

    } catch (error) {
      console.error("Error saving delivery:", error);
      alert("אירעה שגיאה בשמירה");
    }
  };
  


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



  public render(): React.ReactElement<IAfricaDeliveryNewFormProps> {
    const { showForm, formData } = this.state;
  
    return (
      <div className={styles.africaDeliveryNewForm}>
        <button onClick={this.openForm} className={styles.newButton}>
          משלוח חדש
        </button>
  
        {showForm && (
          <div className={styles.popupOverlay} onClick={this.closeForm}>
            <div className={styles.popupContent} onClick={(e) => e.stopPropagation()}>
              {/* כפתור X לסגירה */}
              <div className={styles.closeIcon} onClick={this.closeForm}>
                &times;
              </div>
  
              <h2 style={{ marginTop: 0 }}>טופס חדש לשליחות</h2>
  
              <div className={styles.formGroup}>
                <label>סוג משלוח:</label><br />
                <select
                  name="deliveryType"
                  value={formData.deliveryType}
                  onChange={this.handleDeliveryTypeChange}
                  className={styles.selectField}
                >
                  <option value="">בחר</option>
                  <option value="מסמכי חוזה">מסמכי חוזה</option>
                  <option value="מסמכים כספיים">מסמכים כספיים</option>
                  <option value="ייפוי כוח">ייפוי כוח</option>
                  <option value="מסמכים משפטיים">מסמכים משפטיים</option>
                  <option value="ציוד">ציוד</option>
                  <option value="אחר">אחר</option>
                </select>
              </div>
  

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

              <div className={styles.formGroup}>
                <label>יעד חיצוני:</label><br />
                <input
                  type="text"
                  name="externalDestination"
                  value={formData.externalDestination}
                  onChange={this.handleInputChange}
                  className={styles.inputField}
                />
              </div>
  
              <div className={styles.formGroup}>
                <label>תוכן החבילה:</label><br />
                <textarea
                  name="packageContent"
                  value={formData.packageContent}
                  onChange={this.handleInputChange}
                  className={styles.textAreaField}
                />
              </div>
  
              <div className={styles.formGroup}>
                <label>הערות:</label><br />
                <textarea
                  name="notes"
                  value={formData.notes}
                  onChange={this.handleInputChange}
                  className={styles.textAreaField}
                />
              </div>
  
              <div className={styles.buttonGroup}>
                <button onClick={this.handleSave} className={styles.saveButton}>
                  שמור
                </button>
                <button onClick={this.closeForm} className={styles.closeButton}>
                  סגור
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  
}
