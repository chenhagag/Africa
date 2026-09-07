import * as React from 'react';
import styles from './ExpenseReportForm.module.scss';
import { IExpenseReportFormProps } from './IExpenseReportFormProps';
import { sp, Web } from '@pnp/sp/presets/all';
import { DatePicker, DayOfWeek } from '@fluentui/react';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MSGraphClient = any;

const HebrewDayPickerStrings = {
  months: ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'],
  shortMonths: ['ינו', 'פבר', 'מרץ', 'אפר', 'מאי', 'יונ', 'יול', 'אוג', 'ספט', 'אוק', 'נוב', 'דצמ'],
  days: ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'],
  shortDays: ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש'],
  goToToday: 'היום',
  prevMonthAriaLabel: 'חודש קודם',
  nextMonthAriaLabel: 'חודש הבא',
  prevYearAriaLabel: 'שנה קודמת',
  nextYearAriaLabel: 'שנה הבאה',
};

const pad2 = (n: number): string => n < 10 ? '0' + n : '' + n;

const formatDateIL = (date?: Date): string => {
  if (!date) return '';
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
};

const dateToISO = (date?: Date): string => {
  if (!date) return '';
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
};

const isoToDate = (iso: string): Date | undefined => {
  if (!iso) return undefined;
  const parts = iso.split('-');
  if (parts.length !== 3) return undefined;
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
};

interface IExpenseItem {
  date: string;
  type: string;
  amount: number;
  remarks: string;
}

interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
}

interface ISavedReport {
  Id: number;
  Created: string;
  TotalAmnt: number;
  Status: string;
  Employee: { Title: string } | null;
}

interface IExpenseReportFormState {
  // Form mode
  mode: 'list' | 'new' | 'edit';
  editingItemId: number | null;

  // Form fields
  employeeName: string;
  employeeId: number | null;
  expenses: IExpenseItem[];
  attachments: File[];
  existingAttachments: string[];

  // People picker
  searchText: string;
  suggestedUsers: IGraphUser[];
  allUsers: IGraphUser[];

  // Signatures
  employeeSignatureData: string;
  managerSignatureData: string;
  hasEmployeeDrawn: boolean;
  hasManagerDrawn: boolean;
  employeeSignatureUrl: string;
  managerSignatureUrl: string;

  // UI state
  isLoading: boolean;
  isSaving: boolean;
  statusMessage: string;
  statusType: 'success' | 'error' | '';

  // My reports
  myReports: ISavedReport[];
  currentPage: number;

  // Current user
  currentUserName: string;
  currentUserEmail: string;
}

const EXPENSE_TYPES = ['חניה', 'דלק', 'נסיעות', 'ארוחות', 'אירוח', 'אחר'];


export default class ExpenseReportForm extends React.Component<IExpenseReportFormProps, IExpenseReportFormState> {

  private fileInputRef: React.RefObject<HTMLInputElement>;
  private employeeCanvasRef: React.RefObject<HTMLCanvasElement>;
  private managerCanvasRef: React.RefObject<HTMLCanvasElement>;
  private isDrawing: boolean = false;
  private activeCanvas: 'employee' | 'manager' | null = null;

  constructor(props: IExpenseReportFormProps) {
    super(props);

    const today = new Date().toISOString().split('T')[0];

    this.state = {
      mode: 'list',
      editingItemId: null,
      employeeName: '',
      employeeId: null,
      expenses: [{ date: today, type: 'חניה', amount: 0, remarks: '' }],
      attachments: [],
      existingAttachments: [],
      searchText: '',
      suggestedUsers: [],
      allUsers: [],
      isLoading: true,
      isSaving: false,
      statusMessage: '',
      statusType: '',
      myReports: [],
      currentPage: 1,
      currentUserName: '',
      currentUserEmail: '',
      employeeSignatureData: '',
      managerSignatureData: '',
      hasEmployeeDrawn: false,
      hasManagerDrawn: false,
      employeeSignatureUrl: '',
      managerSignatureUrl: ''
    };

    this.fileInputRef = React.createRef<HTMLInputElement>();
    this.employeeCanvasRef = React.createRef<HTMLCanvasElement>();
    this.managerCanvasRef = React.createRef<HTMLCanvasElement>();
  }

  private get web(): ReturnType<typeof Web> {
    return this.props.siteUrl ? Web(this.props.siteUrl) : sp.web;
  }

  public async componentDidMount(): Promise<void> {
    try {
      // Load current user
      const currentUser = await sp.web.currentUser.get();
      this.setState({
        currentUserName: currentUser.Title,
        currentUserEmail: currentUser.Email
      });

      // Load users from Graph and reports in parallel
      await Promise.all([
        this.loadAllUsers(),
        this.loadMyReports()
      ]);
    } catch (error) {
      console.error('Error initializing:', error);
      this.setState({ statusMessage: 'שגיאה בטעינת הנתונים', statusType: 'error' });
    } finally {
      this.setState({ isLoading: false });
    }
  }

  private async loadAllUsers(): Promise<void> {
    try {
      const client: MSGraphClient = await this.props.context.msGraphClientFactory.getClient('3');
      let allUsers: IGraphUser[] = [];
      let nextLink: string | undefined = '/users?$select=id,displayName,mail,userPrincipalName,assignedLicenses&$top=999';

      while (nextLink) {
        const response: any = await client.api(nextLink).version('v1.0').get();
        allUsers = allUsers.concat(response.value);
        nextLink = response['@odata.nextLink']
          ? response['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
          : undefined;
      }

      const licensedUsers = allUsers.filter((user: any) =>
        user.assignedLicenses && user.assignedLicenses.length > 0
      );

      this.setState({ allUsers: licensedUsers });
    } catch (error) {
      console.error('Error loading users:', error);
    }
  }

  private async loadMyReports(): Promise<void> {
    try {
      // Load reports created by me OR submitted for me (Employee field)
      const currentUser = await sp.web.currentUser.get();
      const userId = currentUser.Id;

      const items = await this.web.lists.getByTitle(this.props.listName).items
        .select('Id', 'Created', 'TotalAmnt', 'Status', 'Employee/Title', 'Author/Title')
        .expand('Employee', 'Author')
        .filter(`(Author/Id eq ${userId}) or (Employee/Id eq ${userId})`)
        .orderBy('Created', false)
        .top(500)();

      this.setState({ myReports: items as ISavedReport[] });
    } catch (error) {
      console.error('Error loading reports:', error);
      // List might not exist yet, that's ok
      this.setState({ myReports: [] });
    }
  }

  private handleUserInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const value = event.target.value;
    this.setState({ searchText: value, employeeName: value });

    if (!value || value.length < 2) {
      this.setState({ suggestedUsers: [] });
      return;
    }

    const filtered = this.state.allUsers.filter(user => {
      const s = value.toLowerCase();
      return (
        (user.displayName && user.displayName.toLowerCase().includes(s)) ||
        (user.mail && user.mail.toLowerCase().includes(s))
      );
    });

    this.setState({ suggestedUsers: filtered.slice(0, 10) });
  }

  private handleUserSelect = async (user: IGraphUser): Promise<void> => {
    try {
      const ensuredUser = await sp.web.ensureUser(user.userPrincipalName);
      this.setState({
        employeeName: user.displayName,
        employeeId: ensuredUser.data.Id,
        searchText: user.displayName,
        suggestedUsers: []
      });
    } catch (error) {
      console.error('Error ensuring user:', error);
      this.setState({
        employeeName: user.displayName,
        searchText: user.displayName,
        suggestedUsers: [],
        statusMessage: 'שגיאה בזיהוי המשתמש',
        statusType: 'error'
      });
    }
  }

  private handleExpenseChange = (index: number, field: keyof IExpenseItem, value: string | number): void => {
    const expenses = [...this.state.expenses];
    (expenses[index] as any)[field] = value;
    this.setState({ expenses });
  }

  private addExpenseRow = (): void => {
    const today = new Date().toISOString().split('T')[0];
    this.setState({
      expenses: [...this.state.expenses, { date: today, type: 'חניה', amount: 0, remarks: '' }]
    });
  }

  private removeExpenseRow = (index: number): void => {
    if (this.state.expenses.length <= 1) return;
    const expenses = [...this.state.expenses];
    expenses.splice(index, 1);
    this.setState({ expenses });
  }

  private calculateTotals(): { [key: string]: number; total: number } {
    const totals: { [key: string]: number; total: number } = { total: 0 };
    EXPENSE_TYPES.forEach(t => { totals[t] = 0; });

    this.state.expenses.forEach(exp => {
      const amt = Number(exp.amount) || 0;
      if (totals[exp.type] !== undefined) {
        totals[exp.type] += amt;
      }
      totals.total += amt;
    });

    return totals;
  }

  private handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>): void => {
    if (event.target.files) {
      const newFiles = Array.from(event.target.files);
      this.setState({ attachments: [...this.state.attachments, ...newFiles] });
      event.target.value = '';
    }
  }

  private removeAttachment = (index: number): void => {
    const attachments = [...this.state.attachments];
    attachments.splice(index, 1);
    this.setState({ attachments });
  }

  private removeExistingAttachment = async (fileName: string): Promise<void> => {
    if (!this.state.editingItemId) return;
    try {
      await this.web.lists.getByTitle(this.props.listName)
        .items.getById(this.state.editingItemId)
        .attachmentFiles.getByName(fileName).delete();

      this.setState({
        existingAttachments: this.state.existingAttachments.filter(f => f !== fileName)
      });
    } catch (error) {
      console.error('Error removing attachment:', error);
    }
  }

  private saveReport = async (status: string): Promise<void> => {
    const { employeeName, employeeId, expenses, attachments, editingItemId } = this.state;

    if (!employeeName || !employeeId) {
      this.setState({ statusMessage: 'יש לבחור עובד', statusType: 'error' });
      return;
    }

    const validExpenses = expenses.filter(e => e.amount > 0);
    if (validExpenses.length === 0) {
      this.setState({ statusMessage: 'יש להזין לפחות הוצאה אחת', statusType: 'error' });
      return;
    }

    this.setState({ isSaving: true, statusMessage: '', statusType: '' });

    try {
      const totals = this.calculateTotals();

      const itemData: any = {
        Title: `החזר הוצאות - ${employeeName}`,
        EmployeeId: employeeId,
        ExpenseItems: JSON.stringify(validExpenses),
        TotalPark: totals['חניה'] || 0,
        TotalFuel: totals['דלק'] || 0,
        TotalTravel: totals['נסיעות'] || 0,
        TotalMeals: totals['ארוחות'] || 0,
        TotalHost: totals['אירוח'] || 0,
        TotalOther: totals['אחר'] || 0,
        TotalAmnt: totals.total,
        Status: status
      };

      let itemId: number;

      if (editingItemId) {
        await this.web.lists.getByTitle(this.props.listName)
          .items.getById(editingItemId).update(itemData);
        itemId = editingItemId;
      } else {
        const result = await this.web.lists.getByTitle(this.props.listName)
          .items.add(itemData);
        itemId = result.data.Id;
      }

      // Upload attachments
      for (const file of attachments) {
        try {
          const arrayBuffer = await file.arrayBuffer();
          await this.web.lists.getByTitle(this.props.listName)
            .items.getById(itemId)
            .attachmentFiles.add(file.name, arrayBuffer);
        } catch (err) {
          console.error(`Error uploading ${file.name}:`, err);
        }
      }

      // Upload signatures (delete old ones first if re-signing)
      if (this.state.hasEmployeeDrawn) {
        const blob = this.getSignatureBlob('employee');
        if (blob) {
          try {
            // Delete old signature if exists
            try {
              await this.web.lists.getByTitle(this.props.listName)
                .items.getById(itemId)
                .attachmentFiles.getByName(`employee-signature-${itemId}.png`).delete();
            } catch { /* no old signature */ }
            const buf = await blob.arrayBuffer();
            await this.web.lists.getByTitle(this.props.listName)
              .items.getById(itemId)
              .attachmentFiles.add(`employee-signature-${itemId}.png`, buf);
          } catch (err) {
            console.error('Error uploading employee signature:', err);
          }
        }
      }
      if (this.state.hasManagerDrawn) {
        const blob = this.getSignatureBlob('manager');
        if (blob) {
          try {
            try {
              await this.web.lists.getByTitle(this.props.listName)
                .items.getById(itemId)
                .attachmentFiles.getByName(`manager-signature-${itemId}.png`).delete();
            } catch { /* no old signature */ }
            const buf = await blob.arrayBuffer();
            await this.web.lists.getByTitle(this.props.listName)
              .items.getById(itemId)
              .attachmentFiles.add(`manager-signature-${itemId}.png`, buf);
          } catch (err) {
            console.error('Error uploading manager signature:', err);
          }
        }
      }

      const statusText = status === 'טיוטה' ? 'הטיוטה נשמרה בהצלחה' : 'הטופס הוגש בהצלחה';
      this.setState({
        statusMessage: statusText,
        statusType: 'success',
        isSaving: false,
        attachments: []
      });

      // Reload reports and go back to list
      await this.loadMyReports();
      setTimeout(() => {
        this.setState({ mode: 'list', statusMessage: '', statusType: '' });
      }, 2000);

    } catch (error) {
      console.error('Error saving report:', error);
      this.setState({
        statusMessage: 'שגיאה בשמירת הטופס: ' + (error as Error).message,
        statusType: 'error',
        isSaving: false
      });
    }
  }

  private loadReport = async (itemId: number): Promise<void> => {
    this.setState({ isLoading: true });

    try {
      const item = await this.web.lists.getByTitle(this.props.listName)
        .items.getById(itemId)
        .select('Id', 'Created', 'ExpenseItems', 'Status', 'Employee/Title', 'Employee/Id', 'Employee/EMail')
        .expand('Employee')();

      const expenses: IExpenseItem[] = item.ExpenseItems ? JSON.parse(item.ExpenseItems) : [];
      if (expenses.length === 0) {
        const today = new Date().toISOString().split('T')[0];
        expenses.push({ date: today, type: 'חניה', amount: 0, remarks: '' });
      }

      // Load existing attachments
      let existingAttachments: string[] = [];
      let employeeSignatureUrl = '';
      let managerSignatureUrl = '';
      try {
        const attachFiles = await this.web.lists.getByTitle(this.props.listName)
          .items.getById(itemId)
          .attachmentFiles();

        const siteUrl = this.props.siteUrl || this.props.context.pageContext.web.absoluteUrl;

        for (const f of attachFiles) {
          const fileName: string = f.FileName;
          if (fileName.indexOf('employee-signature') > -1) {
            employeeSignatureUrl = `${siteUrl}/Lists/${this.props.listName}/Attachments/${itemId}/${fileName}`;
          } else if (fileName.indexOf('manager-signature') > -1) {
            managerSignatureUrl = `${siteUrl}/Lists/${this.props.listName}/Attachments/${itemId}/${fileName}`;
          } else {
            existingAttachments.push(fileName);
          }
        }
      } catch {
        // no attachments
      }

      this.setState({
        mode: 'edit',
        editingItemId: itemId,
        employeeName: item.Employee?.Title || '',
        employeeId: item.Employee?.Id || null,
        searchText: item.Employee?.Title || '',
        expenses,
        existingAttachments,
        attachments: [],
        employeeSignatureUrl,
        managerSignatureUrl,
        hasEmployeeDrawn: false,
        hasManagerDrawn: false,
        isLoading: false
      });
    } catch (error) {
      console.error('Error loading report:', error);
      this.setState({
        isLoading: false,
        statusMessage: 'שגיאה בטעינת הדוח',
        statusType: 'error'
      });
    }
  }

  private startNewReport = async (): Promise<void> => {
    const today = new Date().toISOString().split('T')[0];

    // Default to current user
    let employeeId: number | null = null;
    try {
      const currentUser = await sp.web.currentUser.get();
      employeeId = currentUser.Id;
    } catch {
      // ignore
    }

    this.setState({
      mode: 'new',
      editingItemId: null,
      employeeName: this.state.currentUserName,
      employeeId: employeeId,
      searchText: this.state.currentUserName,
      expenses: [{ date: today, type: 'חניה', amount: 0, remarks: '' }],
      attachments: [],
      existingAttachments: [],
      statusMessage: '',
      statusType: '',
      employeeSignatureData: '',
      managerSignatureData: '',
      hasEmployeeDrawn: false,
      hasManagerDrawn: false,
      employeeSignatureUrl: '',
      managerSignatureUrl: ''
    });
  }

  // --- Signature canvas methods ---

  private getCanvasRef(type: 'employee' | 'manager'): React.RefObject<HTMLCanvasElement> {
    return type === 'employee' ? this.employeeCanvasRef : this.managerCanvasRef;
  }

  private handleCanvasMouseDown = (type: 'employee' | 'manager', e: React.MouseEvent<HTMLCanvasElement>): void => {
    this.activeCanvas = type;
    this.isDrawing = true;
    const canvas = this.getCanvasRef(type).current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.strokeStyle = '#000';
    ctx.beginPath();
    ctx.moveTo(e.nativeEvent.offsetX, e.nativeEvent.offsetY);
    if (type === 'employee') this.setState({ hasEmployeeDrawn: true });
    else this.setState({ hasManagerDrawn: true });
  }

  private handleCanvasTouchStart = (type: 'employee' | 'manager', e: React.TouchEvent<HTMLCanvasElement>): void => {
    e.preventDefault();
    this.activeCanvas = type;
    this.isDrawing = true;
    const canvas = this.getCanvasRef(type).current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    const touch = e.touches[0];
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.strokeStyle = '#000';
    ctx.beginPath();
    ctx.moveTo(touch.clientX - rect.left, touch.clientY - rect.top);
    if (type === 'employee') this.setState({ hasEmployeeDrawn: true });
    else this.setState({ hasManagerDrawn: true });
  }

  private handleCanvasMouseMove = (type: 'employee' | 'manager', e: React.MouseEvent<HTMLCanvasElement>): void => {
    if (!this.isDrawing || this.activeCanvas !== type) return;
    const canvas = this.getCanvasRef(type).current;
    const ctx = canvas?.getContext('2d');
    if (!ctx) return;
    ctx.lineTo(e.nativeEvent.offsetX, e.nativeEvent.offsetY);
    ctx.stroke();
  }

  private handleCanvasTouchMove = (type: 'employee' | 'manager', e: React.TouchEvent<HTMLCanvasElement>): void => {
    e.preventDefault();
    if (!this.isDrawing || this.activeCanvas !== type) return;
    const canvas = this.getCanvasRef(type).current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    const touch = e.touches[0];
    ctx.lineTo(touch.clientX - rect.left, touch.clientY - rect.top);
    ctx.stroke();
  }

  private handleCanvasEnd = (): void => {
    this.isDrawing = false;
    this.activeCanvas = null;
  }

  private clearSignature = (type: 'employee' | 'manager'): void => {
    const canvas = this.getCanvasRef(type).current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    if (type === 'employee') {
      this.setState({ hasEmployeeDrawn: false, employeeSignatureData: '' });
    } else {
      this.setState({ hasManagerDrawn: false, managerSignatureData: '' });
    }
  }

  private dataURLtoBlob(dataUrl: string): Blob {
    const arr = dataUrl.split(',');
    const mimeMatch = arr[0].match(/:(.*?);/);
    const mime = mimeMatch ? mimeMatch[1] : 'image/png';
    const bstr = atob(arr[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) {
      u8arr[n] = bstr.charCodeAt(n);
    }
    return new Blob([u8arr], { type: mime });
  }

  private getSignatureBlob(type: 'employee' | 'manager'): Blob | null {
    const canvas = this.getCanvasRef(type).current;
    if (!canvas) return null;
    const dataUrl = canvas.toDataURL('image/png');
    return this.dataURLtoBlob(dataUrl);
  }

  private handlePrint = (): void => {
    window.print();
  }

  public render(): React.ReactElement<IExpenseReportFormProps> {
    const { mode, isLoading, isSaving, statusMessage, statusType } = this.state;

    if (isLoading) {
      return (
        <div className={styles.expenseReportForm}>
          <div className={styles.loadingOverlay}>טוען...</div>
        </div>
      );
    }

    return (
      <div className={styles.expenseReportForm}>
        <div className={styles.header}>
          <h1>טופס החזר הוצאות אישיות</h1>
        </div>

        {statusMessage && (
          <div className={`${styles.statusBar} ${statusType === 'success' ? styles.success : styles.error}`}>
            {statusMessage}
          </div>
        )}

        {isSaving && <div className={styles.loadingOverlay}>שומר...</div>}

        {mode === 'list' ? this.renderReportsList() : (
          <>
            {this.renderForm()}
            <hr style={{ margin: '32px 0', border: 'none', borderTop: '2px solid #e1dfdd' }} />
            {this.renderReportsList()}
          </>
        )}
      </div>
    );
  }

  private renderReportsList(): React.ReactElement {
    const { myReports, currentPage } = this.state;
    const pageSize = 10;
    const totalPages = Math.ceil(myReports.length / pageSize);
    const startIndex = (currentPage - 1) * pageSize;
    const pageReports = myReports.slice(startIndex, startIndex + pageSize);

    return (
      <div className={styles.myReportsList}>
        <button className={styles.newReportBtn} onClick={this.startNewReport}>
          + דוח חדש
        </button>
        <h3>הדוחות שלי ({myReports.length})</h3>
        {myReports.length === 0 && <p>אין דוחות עדיין</p>}
        {pageReports.map(report => (
          <div
            key={report.Id}
            className={styles.reportItem}
            onClick={() => this.loadReport(report.Id)}
          >
            <span className={styles.reportDate}>
              {report.Created ? new Date(report.Created).toLocaleDateString('he-IL') : '-'}
            </span>
            <span>{report.Employee?.Title || ''}</span>
            <span className={styles.reportTotal}>
              {report.TotalAmnt ? `${report.TotalAmnt.toLocaleString()} ₪` : '0 ₪'}
            </span>
            <span className={`${styles.reportStatus} ${report.Status === 'טיוטה' ? styles.draft : styles.submitted}`}>
              {report.Status || 'טיוטה'}
            </span>
          </div>
        ))}
        {totalPages > 1 && (
          <div className={styles.pagination}>
            <button
              disabled={currentPage <= 1}
              onClick={() => this.setState({ currentPage: currentPage - 1 })}
            >
              הקודם
            </button>
            <span>עמוד {currentPage} מתוך {totalPages}</span>
            <button
              disabled={currentPage >= totalPages}
              onClick={() => this.setState({ currentPage: currentPage + 1 })}
            >
              הבא
            </button>
          </div>
        )}
      </div>
    );
  }

  private renderForm(): React.ReactElement {
    const { searchText, suggestedUsers, expenses, attachments, existingAttachments } = this.state;
    const totals = this.calculateTotals();

    return (
      <>
        {/* Employee */}
        <div className={styles.formInfo}>
          <div className={styles.formGroup}>
            <label>שם העובד</label>
            <input
              type="text"
              value={searchText}
              onChange={this.handleUserInputChange}
              placeholder="הקלד שם עובד..."
            />
            {suggestedUsers.length > 0 && (
              <ul className={styles.userDropdown}>
                {suggestedUsers.map(user => (
                  <li key={user.id} onClick={() => this.handleUserSelect(user)}>
                    {user.displayName} ({user.mail})
                  </li>
                ))}
              </ul>
            )}
          </div>
          <div className={styles.formGroup}>
            <label>תאריך</label>
            <input
              type="text"
              value={formatDateIL(new Date())}
              readOnly
              style={{ backgroundColor: '#f3f2f1', cursor: 'default' }}
            />
          </div>
        </div>

        {/* Expenses Table */}
        <h2 className={styles.sectionTitle}>פירוט ההוצאות</h2>
        <table className={styles.expenseTable}>
          <thead>
            <tr>
              <th>תאריך</th>
              <th>סוג ההוצאה</th>
              <th>סכום (₪)</th>
              <th>הערות</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {expenses.map((exp, idx) => (
              <tr key={idx}>
                <td>
                  <DatePicker
                    value={isoToDate(exp.date)}
                    onSelectDate={(date) => this.handleExpenseChange(idx, 'date', dateToISO(date || undefined))}
                    formatDate={formatDateIL}
                    strings={HebrewDayPickerStrings}
                    firstDayOfWeek={DayOfWeek.Sunday}
                    placeholder="בחר תאריך"
                    style={{ minWidth: 130 }}
                  />
                </td>
                <td>
                  <select
                    value={exp.type}
                    onChange={(e) => this.handleExpenseChange(idx, 'type', e.target.value)}
                  >
                    {EXPENSE_TYPES.map(t => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </td>
                <td>
                  <input
                    type="number"
                    value={exp.amount || ''}
                    min="0"
                    onChange={(e) => this.handleExpenseChange(idx, 'amount', parseFloat(e.target.value) || 0)}
                  />
                </td>
                <td>
                  <input
                    type="text"
                    value={exp.remarks}
                    onChange={(e) => this.handleExpenseChange(idx, 'remarks', e.target.value)}
                    placeholder="הערות..."
                  />
                </td>
                <td>
                  <button
                    className={styles.deleteBtn}
                    onClick={() => this.removeExpenseRow(idx)}
                    title="מחק שורה"
                  >
                    ✕
                  </button>
                </td>
              </tr>
            ))}
            <tr className={styles.totalRow}>
              <td colSpan={2}>סה&quot;כ</td>
              <td>{totals.total.toLocaleString()} ₪</td>
              <td colSpan={2}></td>
            </tr>
          </tbody>
        </table>

        <button className={styles.addBtn} onClick={this.addExpenseRow}>
          + הוסף הוצאה
        </button>

        {/* Attachments */}
        <div className={styles.attachmentsSection}>
          <h2 className={styles.sectionTitle}>קבצים מצורפים (קבלות)</h2>

          {existingAttachments.length > 0 && (
            <div className={styles.attachmentsList}>
              <strong>קבצים קיימים:</strong>
              {existingAttachments.map((fileName, idx) => (
                <div key={idx} className={styles.attachmentItem}>
                  <span>{fileName}</span>
                  <button
                    className={styles.removeAttachment}
                    onClick={() => this.removeExistingAttachment(fileName)}
                  >
                    הסר
                  </button>
                </div>
              ))}
            </div>
          )}

          {attachments.length > 0 && (
            <div className={styles.attachmentsList}>
              <strong>קבצים חדשים:</strong>
              {attachments.map((file, idx) => (
                <div key={idx} className={styles.attachmentItem}>
                  <span>{file.name}</span>
                  <button
                    className={styles.removeAttachment}
                    onClick={() => this.removeAttachment(idx)}
                  >
                    הסר
                  </button>
                </div>
              ))}
            </div>
          )}

          <input
            ref={this.fileInputRef}
            type="file"
            multiple
            onChange={this.handleFileSelect}
          />
        </div>

        {/* Summary by Category */}
        <div className={styles.summarySection}>
          {EXPENSE_TYPES.map(type => (
            <div key={type} className={styles.summaryItem}>
              <span className={styles.summaryLabel}>סה&quot;כ {type}</span>
              <span className={styles.summaryValue}>{(totals[type] || 0).toLocaleString()} ₪</span>
            </div>
          ))}
          <div className={styles.summaryTotal}>
            <span>סה&quot;כ</span>
            <span>{totals.total.toLocaleString()} ₪</span>
          </div>
        </div>

        {/* Signatures */}
        <div className={styles.signatureSection}>
          <div className={styles.signatureBox}>
            <label>חתימת עובד</label>
            {this.state.employeeSignatureUrl && !this.state.hasEmployeeDrawn ? (
              <div>
                <img src={this.state.employeeSignatureUrl} alt="חתימת עובד" className={styles.signatureImage} />
                <button
                  className={styles.clearSignatureBtn}
                  onClick={() => this.setState({ employeeSignatureUrl: '' })}
                >
                  חתום מחדש
                </button>
              </div>
            ) : (
              <div>
                <canvas
                  ref={this.employeeCanvasRef}
                  width={300}
                  height={150}
                  className={styles.signatureCanvas}
                  onMouseDown={(e) => this.handleCanvasMouseDown('employee', e)}
                  onMouseMove={(e) => this.handleCanvasMouseMove('employee', e)}
                  onMouseUp={this.handleCanvasEnd}
                  onMouseLeave={this.handleCanvasEnd}
                  onTouchStart={(e) => this.handleCanvasTouchStart('employee', e)}
                  onTouchMove={(e) => this.handleCanvasTouchMove('employee', e)}
                  onTouchEnd={this.handleCanvasEnd}
                  style={{ touchAction: 'none' }}
                />
                <button
                  className={styles.clearSignatureBtn}
                  onClick={() => this.clearSignature('employee')}
                >
                  נקה חתימה
                </button>
              </div>
            )}
          </div>
          <div className={styles.signatureBox}>
            <label>חתימת מנהל</label>
            {this.state.managerSignatureUrl && !this.state.hasManagerDrawn ? (
              <div>
                <img src={this.state.managerSignatureUrl} alt="חתימת מנהל" className={styles.signatureImage} />
                <button
                  className={styles.clearSignatureBtn}
                  onClick={() => this.setState({ managerSignatureUrl: '' })}
                >
                  חתום מחדש
                </button>
              </div>
            ) : (
              <div>
                <canvas
                  ref={this.managerCanvasRef}
                  width={300}
                  height={150}
                  className={styles.signatureCanvas}
                  onMouseDown={(e) => this.handleCanvasMouseDown('manager', e)}
                  onMouseMove={(e) => this.handleCanvasMouseMove('manager', e)}
                  onMouseUp={this.handleCanvasEnd}
                  onMouseLeave={this.handleCanvasEnd}
                  onTouchStart={(e) => this.handleCanvasTouchStart('manager', e)}
                  onTouchMove={(e) => this.handleCanvasTouchMove('manager', e)}
                  onTouchEnd={this.handleCanvasEnd}
                  style={{ touchAction: 'none' }}
                />
                <button
                  className={styles.clearSignatureBtn}
                  onClick={() => this.clearSignature('manager')}
                >
                  נקה חתימה
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Action Buttons */}
        <div className={styles.actions}>
          <button
            className={styles.saveBtn}
            onClick={() => this.saveReport('טיוטה')}
            disabled={this.state.isSaving}
          >
            שמור טיוטה
          </button>
          <button
            className={styles.submitBtn}
            onClick={() => this.saveReport('הוגש')}
            disabled={this.state.isSaving}
          >
            הגש
          </button>
          <button
            className={styles.printBtn}
            onClick={this.handlePrint}
          >
            הדפס
          </button>
          <button
            className={styles.saveBtn}
            onClick={() => this.setState({ mode: 'list', statusMessage: '', statusType: '' })}
            disabled={this.state.isSaving}
          >
            חזור לרשימה
          </button>
        </div>
      </>
    );
  }
}
