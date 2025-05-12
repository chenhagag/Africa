import * as React from 'react';
import { sp } from "@pnp/sp/presets/all";

import { ICrmFileAnalistProps } from './ICrmFileAnalistProps';

interface IParsedLine {
  name: string;
  supplierNum: string;
  address: string;
  city: string;
  zipCode: string; // מיקוד
  mana: string;
  CreditMessage: string;
  creditNoteDate: string;
  invoices: string[];
  general: string;
  bankBranchName: string;
  creationDate: string;
  accountNumber: string;
  referenceNumber: string;
  bankBranch: string;        // CKMSUB
  bankNumber: string;        // CKMBNK
  bankAccount: string;       // CKMCHS
  bankBranchCity: string;       // CKMSGH
  bankName: string;     // CKMSHL
  bankBranchAddress: string;      // CKMATA
  extraText1: string;
  extraText2: string;
  message: string;
}


interface IState {
  lines: IParsedLine[];
  fileUploaded: boolean;
  errorMessage: string | null;
  showConfirmation: boolean;
  isSending: boolean;
}

export default class CrmFileAnalistWebPart extends React.Component<ICrmFileAnalistProps, IState> {
  constructor(props: ICrmFileAnalistProps) {

      super(props);
    this.state = {
      lines: [],
      fileUploaded: false,
      errorMessage: null,
      showConfirmation: false,
      isSending: false
    };
  }

  componentDidMount(): void {
    sp.setup({
      sp: {
        baseUrl: "https://africaisrael.sharepoint.com/IT-Invntory"
      }
    });
  }


  private async loadPositionConfig(): Promise<any> {

    const response = await fetch(`https://africaisrael.sharepoint.com/IT-Invntory/SiteAssets/CRMConfig.JSON`);
    if (!response.ok) {
      throw new Error('Failed to load configuration');
    }
    return response.json();
  }
  
  private handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {

    const file = e.target.files?.[0];
    if (!file) return;
  
    try {
      const text = await file.text();
      const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
  
      const config = await this.loadPositionConfig();
      const parsedLines = lines.map(line => this.parseLineWithConfig(line, config));
  
      const isValid = parsedLines.length > 0 && parsedLines[0].name?.length > 0;
      this.setState({ 
        lines: isValid ? parsedLines : [], 
        fileUploaded: isValid, 
        errorMessage: isValid ? null : 'הקובץ שהועלה אינו תקין או שלא עבר המרה כהלכה.' 
      });
    } catch (err) {
      this.setState({ lines: [], fileUploaded: false, errorMessage: 'אירעה שגיאה בעת קריאת הקובץ.' });
    }
  };

  private parseLineWithConfig(line: string, config: any): IParsedLine {
    const get = (from: number, to: number) => line.slice(from - 1, to).trim();

    const parsed: any = {};
    for (const [key, range] of Object.entries(config) as [string, { start: number; end: number }][]) {
      parsed[key] = get(range.start, range.end);
    }    
  
    const invFields = config.invoices.fields;
    const invoiceCount = config.invoices.count;
    
    parsed.invoices = Array.from({ length: invoiceCount }, (_, i) => {
      const number = get(invFields.number.start + i * invFields.number.length, invFields.number.start + i * invFields.number.length + invFields.number.length - 1);
      if (!number.trim()) return null;
    
      const desc = get(invFields.desc.start + i * invFields.desc.length, invFields.desc.start + i * invFields.desc.length + invFields.desc.length - 1);
      const grossAmount = get(invFields.grossAmount.start + i * invFields.grossAmount.length, invFields.grossAmount.start + i * invFields.grossAmount.length + invFields.grossAmount.length - 1);
      const deductionPercent = get(invFields.deductionPercent.start + i * invFields.deductionPercent.length, invFields.deductionPercent.start + i * invFields.deductionPercent.length + invFields.deductionPercent.length - 1);
      const deductionAmount = get(invFields.deductionAmount.start + i * invFields.deductionAmount.length, invFields.deductionAmount.start + i * invFields.deductionAmount.length + invFields.deductionAmount.length - 1);
      const netAmount = get(invFields.netAmount.start + i * invFields.netAmount.length, invFields.netAmount.start + i * invFields.netAmount.length + invFields.netAmount.length - 1);
    
      return `חשבונית מס ${number} | פרטים: ${desc} | סכום: ${grossAmount} | ניכוי %: ${deductionPercent} | סכום ניכוי: ${deductionAmount} | לתשלום: ${netAmount}`;
    }).filter(Boolean);
    
    
    return parsed as IParsedLine;
  }
  

  private handleSendClick = () => {
    this.setState({ showConfirmation: true });
  };

  private confirmSend = () => {
    this.setState({ showConfirmation: false, isSending: true });
    this.sendData();
  };

  private CreateMailContent(line: any): string {
    
    const name = line.name || '';
    const idNumber = line.referenceNumber || '';
    const invoices = line.invoices || [];
  
    // חשבוניות כתצוגת שורות רגילות
    const invoiceSection = invoices.length > 0
      ? invoices.map((inv: string) => `<div>${inv}</div>`).join('')
      : `<div>חשבונית מס/עסקה: טרם התקבלה</div>`;
  
    // סכום כולל לתשלום
    const totalToPay = invoices.reduce((sum: number, inv: string) => {
      const match = inv.match(/לתשלום: ([\d.]+)/);
      if (match && match[1]) {
        return sum + parseFloat(match[1]);
      }
      return sum;
    }, 0).toFixed(2);
  
    // הצגת הטקסט של "יש להפיק חשבוניות..." רק אם אין חשבוניות
    const missingInvoiceNote = invoices.length === 0
      ? `
        <p style="margin-top: 1em;">
          את החשבוניות יש להפיק עבור חברת <strong>"__________"</strong>, ח.פ. <strong>__________</strong>
        </p>`
      : '';
  
    const fullHtml = `
      <div dir="rtl" style="font-family: Arial; font-size: 14px;">
        <p>לכבוד,</p>
        <p><strong>${name}</strong></p>
        <p>ח.פ./עוסק מורשה <strong>${idNumber}</strong></p>
  
        <p style="margin-top: 1em;">
          הרינו להודיעך שזיכינו את חשבונך בבנק בהתאם לפירוט הוראות התשלום הבאות:
        </p>
  
        ${invoiceSection}
  
        ${missingInvoiceNote}
  
        <p style="margin-top: 1em;"><strong>סה"כ לתשלום: ${totalToPay}</strong></p>
  
        <p style="margin-top: 1em;"><strong>פרטי חשבון בנק:</strong></p>
        <p>בנק: ${line.bankName}</p>
        <p>סניף: ${line.bankBranchName} - ${line.bankBranch}</p>
        <p>חשבון: ${line.accountNumber}</p>
  
        <p style="margin-top: 1.5em;">${line.extraText1 || ''}</p>
  
        <p style="margin-top: 1.5em;">בברכה,<br/>אפריקה ישראל מגורים</p>
      </div>
    `;
  
    return fullHtml;
  }
  
  

  private sendData = async () => {
    const fileName = (document.getElementById('upload') as HTMLInputElement)?.files?.[0]?.name ?? 'UnknownFile.txt';
  
    try {
      for (const line of this.state.lines) {
      
        const customerData = [
          `שם: ${line.name}`,
          `כתובת: ${line.address}`,
          `עיר: ${line.city}`,
          `מיקוד: ${line.zipCode}`,
          `תאריך יצירה: ${line.creationDate}`,
          `תעודת זהות: ${line.referenceNumber}`
        ].join('\n');
  
        const bankData = [
          `חשבון: ${line.accountNumber}`,
          `סניף: ${line.bankBranch}`,
          `שם סניף: ${line.bankBranchName}`,
          `מספר בנק: ${line.bankNumber}`,
          `שם בנק: ${line.bankName}`,
          `כתובת סניף: ${line.bankBranchAddress}`,
          `עיר סניף: ${line.bankBranchCity}`
        ].join('\n');
  
        const invoicesText = line.invoices.join('\n');
        const content = this.CreateMailContent(line)
  
        await sp.web.lists.getByTitle("CRM").items.add({
          Title: line.name,
          Mail: line.extraText2,
          CustomerData: customerData,
          BankData: bankData,
          Invoices: invoicesText,
          FileName: fileName,
          mailContent: content,
          message: line.CreditMessage
        });
      }
  
      this.setState({ isSending: false });
      alert("הנתונים נשמרו בהצלחה לרשימה");
    } catch (error) {
      console.error("שגיאה בשמירת הנתונים:", error);
      alert("אירעה שגיאה בשמירת הנתונים. בדוק את הקונסול לפרטים.");
    }
  };
  
  

  public render(): React.ReactElement {
    const { lines, fileUploaded, errorMessage, showConfirmation } = this.state;

    return (
      <div style={{ fontFamily: 'Segoe UI, sans-serif', padding: '1em' }}>
        {!fileUploaded && (
          <div style={{
            backgroundColor: '#fffbe6',
            border: '1px solid #ffe58f',
            borderRadius: '4px',
            padding: '1em',
            marginBottom: '1em',
            lineHeight: '1.6'
          }}>
            <strong>לפני העלאת הקובץ:</strong>
            <ol style={{ margin: 0, paddingInlineStart: '1.2em' }}>
              <li>פתחו את הקובץ ב־<strong>Notepad++</strong>.</li>
              <li>בחרו בתפריט <strong>Encoding → Character Sets → Hebrew → OEM 862</strong>.</li>
              <li>לאחר מכן בחרו <strong>Encoding → Convert to UTF-8</strong>.</li>
              <li>שמרו את הקובץ החדש והעלו אותו כאן.</li>
            </ol>
          </div>
        )}

        {errorMessage && (
          <div style={{
            color: '#a80000',
            backgroundColor: '#fde7e9',
            border: '1px solid #f5c6cb',
            borderRadius: '4px',
            padding: '0.75em',
            marginBottom: '1em'
          }}>
            {errorMessage}
          </div>
        )}

        <div style={{ marginBottom: '1em' }}>
          <label htmlFor="upload" style={{
            display: 'inline-block',
            padding: '0.5em 1em',
            backgroundColor: '#0078d4',
            color: '#fff',
            borderRadius: '4px',
            cursor: 'pointer',
            fontSize: '14px'
          }}>
            העלה קובץ
          </label>
          <input id="upload" type="file" onChange={this.handleFile} style={{ display: 'none' }} />
        </div>

        {lines.length > 0 && (
          <>
            {this.state.isSending ? (
                <div style={{ margin: '1em 0' }}>
                  <span>שולח נתונים...</span>
                  <div className="spinner" style={{
                    display: 'inline-block',
                    width: '20px',
                    height: '20px',
                    border: '3px solid #ccc',
                    borderTop: '3px solid #107c10',
                    borderRadius: '50%',
                    animation: 'spin 1s linear infinite',
                    marginLeft: '10px',
                    verticalAlign: 'middle'
                  }} />
                </div>
              ) : (
                <div style={{ display: 'flex', justifyContent: 'flex-start', marginBottom: '1em' }}>
                  <button
                    onClick={this.handleSendClick}
                    style={{
                      padding: '0.5em 1em',
                      backgroundColor: '#107c10',
                      color: '#fff',
                      border: 'none',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      fontSize: '14px'
                    }}>
                    שלח
                  </button>
                </div>
              )}


            {showConfirmation && (
              <div style={{
                backgroundColor: '#ffffff',
                border: '1px solid #ccc',
                padding: '1em',
                borderRadius: '4px',
                marginBottom: '1em',
                maxWidth: '400px',
                boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
              }}>
                <p style={{ marginBottom: '1em' }}>
                  שליחת הנתונים תשלח מייל לכל אחד מהספקים עם הנתונים שלו. האם ברצונך להתקדם?
                </p>
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5em' }}>
                  <button onClick={this.confirmSend} style={{
                    padding: '0.4em 1em',
                    backgroundColor: '#0078d4',
                    color: '#fff',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: 'pointer'
                  }}>
                    שליחה
                  </button>
                  <button onClick={() => this.setState({ showConfirmation: false })} style={{
                    padding: '0.4em 1em',
                    backgroundColor: '#f3f2f1',
                    color: '#333',
                    border: '1px solid #ccc',
                    borderRadius: '4px',
                    cursor: 'pointer'
                  }}>
                    ביטול
                  </button>
                </div>
              </div>
            )}

            <table style={{
              width: '100%',
              borderCollapse: 'collapse',
              fontSize: '13px',
              direction: 'rtl'
            }}>
              <thead style={{ backgroundColor: 'darkseagreen' }}>
              <tr>
                {[
                  'שם ספק', 'מספר ספק',
                  'תאריך הפקה', 'מנה', 'הודעת זיכוי', 'תאריך הודעה',
                  'חשבון בנק', 'סניף בנק',
                  'שם סניף בנק', 'מספר בנק', 'שם בנק','כתובת סניף בנק','עיר סניף בנק',
                   'פרויקט', 'לבירורים', 'מייל'
                ].map((header, i) => (
                  <th key={i} style={{ border: '1px solid #ccc', padding: '8px', textAlign: 'right' }}>
                    {header}
                  </th>
                ))}
              </tr>

              </thead>
              <tbody>
  {lines.map((line, index) => (
    <React.Fragment key={index}>
      <tr style={{ backgroundColor: index % 2 === 0 ? '#fff' : '#f9f9f9' }}>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.name}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.supplierNum}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.creationDate}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.mana}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.CreditMessage}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.creditNoteDate}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.accountNumber}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.bankBranch}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.bankBranchName}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.bankNumber}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.bankName}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.bankBranchAddress}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.bankBranchCity}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.general}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.extraText1}</td>
        <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.extraText2}</td>
      </tr>

      <tr>
        <td colSpan={18} style={{ border: '1px solid #ddd', padding: '6px', backgroundColor: '#f1f1f1' }}>
          <strong>חשבוניות:</strong>
          <ul style={{ paddingInlineStart: '1.2em', marginTop: '0.5em', marginBottom: 0 }}>
            {line.invoices.map((inv, i) => (
              <li key={i} style={{ marginBottom: '4px' }}>{inv}</li>
            ))}
          </ul>
        </td>
      </tr>
    </React.Fragment>
  ))}
</tbody>

            </table>
          </>
        )}
      </div>
    );
  }
}
