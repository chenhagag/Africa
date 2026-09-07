import * as React from 'react';
import { IExpenseMigrationProps } from './IExpenseMigrationProps';
import { Web } from '@pnp/sp/presets/all';

interface IExpenseItem {
  date: string;
  type: string;
  amount: number;
  remarks: string;
}

interface ILogEntry {
  fileName: string;
  status: 'success' | 'error' | 'skipped';
  message: string;
}

interface IExpenseMigrationState {
  isRunning: boolean;
  isDryRun: boolean;
  log: ILogEntry[];
  totalFiles: number;
  processedFiles: number;
  successCount: number;
  errorCount: number;
  skippedCount: number;
  testItemId: string;
}

const EXPENSE_TYPE_MAP: { [key: string]: string } = {
  'חניה': 'חניה',
  'דלק': 'דלק',
  'נסיעות': 'נסיעות',
  'ארוחות': 'ארוחות',
  'אירוח': 'אירוח',
  'אחר': 'אחר'
};

export default class ExpenseMigration extends React.Component<IExpenseMigrationProps, IExpenseMigrationState> {

  constructor(props: IExpenseMigrationProps) {
    super(props);
    this.state = {
      isRunning: false,
      isDryRun: true,
      log: [],
      totalFiles: 0,
      processedFiles: 0,
      successCount: 0,
      errorCount: 0,
      skippedCount: 0,
      testItemId: ''
    };
  }

  private get sourceWeb(): ReturnType<typeof Web> {
    return Web(this.props.sourceSiteUrl);
  }

  private get targetWeb(): ReturnType<typeof Web> {
    return Web(this.props.targetSiteUrl);
  }

  private addLog(fileName: string, status: 'success' | 'error' | 'skipped', message: string): void {
    this.setState(prev => ({
      log: [...prev.log, { fileName, status, message }],
      processedFiles: prev.processedFiles + 1,
      successCount: prev.successCount + (status === 'success' ? 1 : 0),
      errorCount: prev.errorCount + (status === 'error' ? 1 : 0),
      skippedCount: prev.skippedCount + (status === 'skipped' ? 1 : 0)
    }));
  }

  private parseXml(xmlString: string): { employee: string; formDate: string; expenses: IExpenseItem[] } | null {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xmlString, 'text/xml');

      // Extract employee - claims identity from DisplayName
      const displayNameEl = doc.getElementsByTagName('pc:DisplayName')[0];
      let employee = '';
      if (displayNameEl && displayNameEl.textContent) {
        // Format: i:0#.f|membership|email@domain.com
        const parts = displayNameEl.textContent.split('|');
        employee = parts.length >= 3 ? parts[2] : displayNameEl.textContent;
      }

      // Extract form date
      const formDateEl = doc.getElementsByTagName('my:FormDate')[0];
      const formDate = formDateEl && formDateEl.textContent ? formDateEl.textContent : '';

      // Extract expenses
      const expenses: IExpenseItem[] = [];
      const expenseNodes = doc.getElementsByTagName('my:Expense');
      for (let i = 0; i < expenseNodes.length; i++) {
        const node = expenseNodes[i];
        const dateEl = node.getElementsByTagName('my:ExpenseDate')[0];
        const typeEl = node.getElementsByTagName('my:ExpenseType')[0];
        const amountEl = node.getElementsByTagName('my:ExpnsAmnt')[0];
        const remarksEl = node.getElementsByTagName('my:ExpnsRemarks')[0];

        const expType = typeEl && typeEl.textContent ? typeEl.textContent : 'אחר';

        expenses.push({
          date: dateEl && dateEl.textContent ? dateEl.textContent : '',
          type: EXPENSE_TYPE_MAP[expType] || 'אחר',
          amount: amountEl && amountEl.textContent ? parseFloat(amountEl.textContent) || 0 : 0,
          remarks: remarksEl && remarksEl.textContent ? remarksEl.textContent : ''
        });
      }

      return { employee, formDate, expenses };
    } catch (error) {
      console.error('XML parse error:', error);
      return null;
    }
  }

  private calculateTotals(expenses: IExpenseItem[]): { [key: string]: number } {
    const totals: { [key: string]: number } = {
      TotalPark: 0, TotalFuel: 0, TotalTravel: 0,
      TotalMeals: 0, TotalHost: 0, TotalOther: 0, TotalAmnt: 0
    };

    const typeToField: { [key: string]: string } = {
      'חניה': 'TotalPark', 'דלק': 'TotalFuel', 'נסיעות': 'TotalTravel',
      'ארוחות': 'TotalMeals', 'אירוח': 'TotalHost', 'אחר': 'TotalOther'
    };

    for (const exp of expenses) {
      const amt = exp.amount || 0;
      const field = typeToField[exp.type] || 'TotalOther';
      totals[field] += amt;
      totals.TotalAmnt += amt;
    }

    return totals;
  }

  private runSingleTest = async (): Promise<void> => {
    const itemId = parseInt(this.state.testItemId);
    if (!itemId || isNaN(itemId)) {
      alert('יש להזין ID תקין');
      return;
    }

    this.setState({
      isRunning: true,
      isDryRun: false,
      log: [],
      totalFiles: 1,
      processedFiles: 0,
      successCount: 0,
      errorCount: 0,
      skippedCount: 0
    });

    try {
      const item = await this.sourceWeb.lists.getByTitle(this.props.sourceLibrary).items
        .getById(itemId)
        .select('Id', 'FileRef', 'FileLeafRef')();

      const fileName: string = item.FileLeafRef;
      const fileRef: string = item.FileRef;

      const fileContent = await this.sourceWeb.getFileByServerRelativeUrl(fileRef).getText();

      const parsed = this.parseXml(fileContent);
      if (!parsed) {
        this.addLog(fileName, 'error', 'לא ניתן לפרסר את ה-XML');
        this.setState({ isRunning: false });
        return;
      }

      if (!parsed.employee) {
        this.addLog(fileName, 'error', 'לא נמצא עובד ב-XML');
        this.setState({ isRunning: false });
        return;
      }

      const totals = this.calculateTotals(parsed.expenses);

      // Resolve user
      let employeeId: number | null = null;
      try {
        const ensured = await this.targetWeb.ensureUser(parsed.employee);
        employeeId = ensured.data.Id;
      } catch {
        try {
          const ensured = await this.targetWeb.ensureUser('i:0#.f|membership|' + parsed.employee);
          employeeId = ensured.data.Id;
        } catch {
          this.addLog(fileName, 'error', `לא ניתן לזהות משתמש: ${parsed.employee}`);
          this.setState({ isRunning: false });
          return;
        }
      }

      const itemData: any = {
        Title: `החזר הוצאות - ${parsed.employee}`,
        EmployeeId: employeeId,
        ExpenseItems: JSON.stringify(parsed.expenses),
        TotalPark: totals.TotalPark,
        TotalFuel: totals.TotalFuel,
        TotalTravel: totals.TotalTravel,
        TotalMeals: totals.TotalMeals,
        TotalHost: totals.TotalHost,
        TotalOther: totals.TotalOther,
        TotalAmnt: totals.TotalAmnt,
        Status: 'הוגש'
      };

      const result = await this.targetWeb.lists.getByTitle(this.props.targetList).items.add(itemData);
      this.addLog(fileName, 'success',
        `נוצר פריט #${result.data.Id} | עובד: ${parsed.employee} | ${parsed.expenses.length} הוצאות | ${totals.TotalAmnt} ₪`
      );

    } catch (error) {
      this.addLog(`ID: ${itemId}`, 'error', `שגיאה: ${(error as Error).message}`);
    }

    this.setState({ isRunning: false });
  }

  private runMigration = async (dryRun: boolean): Promise<void> => {
    this.setState({
      isRunning: true,
      isDryRun: dryRun,
      log: [],
      totalFiles: 0,
      processedFiles: 0,
      successCount: 0,
      errorCount: 0,
      skippedCount: 0
    });

    try {
      // Get all XML files from the source library
      const items = await this.sourceWeb.lists.getByTitle(this.props.sourceLibrary).items
        .select('Id', 'FileRef', 'FileLeafRef')
        .filter("FileLeafRef ne 'template.xsn'")
        .top(500)();

      const xmlItems = items.filter((item: any) => {
        const name: string = item.FileLeafRef || '';
        return name.toLowerCase().indexOf('.xml') > -1;
      });

      this.setState({ totalFiles: xmlItems.length });

      for (const item of xmlItems) {
        const fileName: string = item.FileLeafRef;
        const fileRef: string = item.FileRef;

        try {
          // Download XML content
          const fileContent = await this.sourceWeb.getFileByServerRelativeUrl(fileRef).getText();

          // Parse XML
          const parsed = this.parseXml(fileContent);
          if (!parsed) {
            this.addLog(fileName, 'error', 'לא ניתן לפרסר את ה-XML');
            continue;
          }

          if (!parsed.employee) {
            this.addLog(fileName, 'error', 'לא נמצא עובד ב-XML');
            continue;
          }

          if (parsed.expenses.length === 0) {
            this.addLog(fileName, 'skipped', 'אין הוצאות בטופס');
            continue;
          }

          // Calculate totals
          const totals = this.calculateTotals(parsed.expenses);

          if (dryRun) {
            this.addLog(fileName, 'success',
              `עובד: ${parsed.employee} | תאריך: ${parsed.formDate} | ${parsed.expenses.length} הוצאות | סה"כ: ${totals.TotalAmnt} ₪`
            );
            continue;
          }

          // Resolve user
          let employeeId: number | null = null;
          try {
            const ensured = await this.targetWeb.ensureUser(parsed.employee);
            employeeId = ensured.data.Id;
          } catch {
            // Try with full claims identity
            try {
              const ensured = await this.targetWeb.ensureUser('i:0#.f|membership|' + parsed.employee);
              employeeId = ensured.data.Id;
            } catch {
              this.addLog(fileName, 'error', `לא ניתן לזהות משתמש: ${parsed.employee}`);
              continue;
            }
          }

          // Create item in target list
          const itemData: any = {
            Title: `החזר הוצאות - ${parsed.employee}`,
            EmployeeId: employeeId,
            ExpenseItems: JSON.stringify(parsed.expenses),
            TotalPark: totals.TotalPark,
            TotalFuel: totals.TotalFuel,
            TotalTravel: totals.TotalTravel,
            TotalMeals: totals.TotalMeals,
            TotalHost: totals.TotalHost,
            TotalOther: totals.TotalOther,
            TotalAmnt: totals.TotalAmnt,
            Status: 'הוגש'
          };

          await this.targetWeb.lists.getByTitle(this.props.targetList).items.add(itemData);

          this.addLog(fileName, 'success',
            `הועבר בהצלחה | עובד: ${parsed.employee} | ${parsed.expenses.length} הוצאות | ${totals.TotalAmnt} ₪`
          );

          // Small delay to avoid throttling
          await new Promise(resolve => setTimeout(resolve, 200));

        } catch (error) {
          this.addLog(fileName, 'error', `שגיאה: ${(error as Error).message}`);
        }
      }

    } catch (error) {
      this.addLog('כללי', 'error', `שגיאה בטעינת הספרייה: ${(error as Error).message}`);
    }

    this.setState({ isRunning: false });
  }

  public render(): React.ReactElement<IExpenseMigrationProps> {
    const { isRunning, log, totalFiles, processedFiles, successCount, errorCount, skippedCount, isDryRun } = this.state;

    const containerStyle: React.CSSProperties = {
      fontFamily: "'Segoe UI', sans-serif",
      direction: 'rtl',
      textAlign: 'right',
      maxWidth: 900,
      margin: '0 auto',
      padding: 20
    };

    const btnStyle: React.CSSProperties = {
      padding: '10px 24px',
      fontSize: 14,
      borderRadius: 4,
      cursor: isRunning ? 'not-allowed' : 'pointer',
      fontFamily: 'inherit',
      fontWeight: 600,
      border: 'none',
      marginLeft: 12,
      color: 'white'
    };

    const logStyle: React.CSSProperties = {
      maxHeight: 400,
      overflowY: 'auto',
      border: '1px solid #e1dfdd',
      borderRadius: 4,
      padding: 12,
      marginTop: 16,
      fontSize: 13,
      fontFamily: 'Consolas, monospace',
      direction: 'ltr',
      textAlign: 'left'
    };

    return (
      <div style={containerStyle}>
        <h1 style={{ color: '#0078d4' }}>הסבת טפסי החזר הוצאות</h1>

        <div style={{ background: '#faf9f8', padding: 16, borderRadius: 4, marginBottom: 16 }}>
          <p><strong>מקור:</strong> {this.props.sourceSiteUrl}/{this.props.sourceLibrary}</p>
          <p><strong>יעד:</strong> {this.props.targetSiteUrl} / רשימת {this.props.targetList}</p>
        </div>

        <div style={{ marginBottom: 16 }}>
          <button
            style={{ ...btnStyle, background: '#0078d4' }}
            onClick={() => this.runMigration(true)}
            disabled={isRunning}
          >
            הרצת ניסיון (Dry Run)
          </button>
          <button
            style={{ ...btnStyle, background: '#107c10' }}
            onClick={() => {
              if (confirm('האם להריץ את ההסבה? פעולה זו תיצור פריטים ברשימת היעד.')) {
                this.runMigration(false);
              }
            }}
            disabled={isRunning}
          >
            הרץ הסבה
          </button>
        </div>

        <div style={{ marginBottom: 16, padding: 16, background: '#fff4ce', borderRadius: 4, border: '1px solid #ffb900' }}>
          <strong>בדיקת פריט בודד:</strong>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 8 }}>
            <input
              type="number"
              placeholder="הזן ID של פריט"
              value={this.state.testItemId}
              onChange={(e) => this.setState({ testItemId: e.target.value })}
              style={{ padding: '8px 10px', fontSize: 14, border: '1px solid #8a8886', borderRadius: 4, width: 150, fontFamily: 'inherit' }}
            />
            <button
              style={{ ...btnStyle, background: '#ff8c00' }}
              onClick={this.runSingleTest}
              disabled={isRunning}
            >
              הרץ על פריט בודד
            </button>
          </div>
        </div>

        {isRunning && (
          <div style={{ padding: 12, background: '#deecf9', borderRadius: 4, marginBottom: 12 }}>
            {isDryRun ? 'מריץ בדיקה' : 'מבצע הסבה'}... {processedFiles} / {totalFiles}
          </div>
        )}

        {!isRunning && processedFiles > 0 && (
          <div style={{ padding: 12, background: '#f3f2f1', borderRadius: 4, marginBottom: 12 }}>
            <strong>סיכום{isDryRun ? ' (ניסיון)' : ''}:</strong> {successCount} הצליחו | {errorCount} שגיאות | {skippedCount} דולגו | מתוך {totalFiles} קבצים
          </div>
        )}

        {log.length > 0 && (
          <div style={logStyle}>
            {log.map((entry, idx) => {
              const color = entry.status === 'success' ? '#107c10' : entry.status === 'error' ? '#a4262c' : '#7a6400';
              return (
                <div key={idx} style={{ marginBottom: 4 }}>
                  <span style={{ color }}>[{entry.status === 'success' ? 'OK' : entry.status === 'error' ? 'ERR' : 'SKIP'}]</span>{' '}
                  <strong>{entry.fileName}</strong> — {entry.message}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }
}
