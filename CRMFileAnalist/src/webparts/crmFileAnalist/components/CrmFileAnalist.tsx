import * as React from 'react';
import { ICrmFileAnalistProps } from './ICrmFileAnalistProps';

interface IParsedLine {
  name: string;
  address: string;
  city: string;
  creditNoteNumber: string;
  creditNoteDate: string;
  invoices: string[];
  general: string;
  telephone: string;
  creationDate: string;
  accountNumber: string;
  idNumber: string;
  amount: string;
  referenceNumber: string;
  extraText1: string;
  extraText2: string;
}

interface IState {
  lines: IParsedLine[];
  fileUploaded: boolean;
  errorMessage: string | null;
}

export default class CrmFileAnalistWebPart extends React.Component<ICrmFileAnalistProps, IState> {
  constructor(props: ICrmFileAnalistProps) {
    super(props);
    this.state = {
      lines: [],
      fileUploaded: false,
      errorMessage: null
    };
  }

  private parseLine = (line: string): IParsedLine => {
    const get = (from: number, to: number) => line.slice(from - 1, to).trim();

    const invoices = Array.from({ length: 39 }, (_, i) => {
      const number = get(691 + i * 6, 691 + i * 6 + 5);
      if (!number.trim()) return null;

      const desc = get(925 + i * 25, 925 + i * 25 + 24);
      const grossAmount = get(1900 + i * 15, 1900 + i * 15 + 14);
      const deductionPercent = get(2485 + i * 5, 2485 + i * 5 + 4);
      const deductionAmount = get(2680 + i * 15, 2680 + i * 15 + 14);
      const netAmount = get(3265 + i * 15, 3265 + i * 15 + 14);

      return `חשבונית מס ${number}, פרטים: ${desc}, סכום: ${grossAmount}, ניכוי %: ${deductionPercent}, סכום ניכוי: ${deductionAmount}, לתשלום: ${netAmount}`;
    }).filter(Boolean) as string[];

    return {
      name: get(2, 26),
      address: get(27, 46),
      city: get(47, 61),
      creditNoteNumber: get(82, 88),
      creditNoteDate: get(89, 96),
      accountNumber: get(97, 113),
      idNumber: get(114, 123),
      telephone: get(124, 133),
      creationDate: get(67, 74),
      amount: get(177, 187),
      referenceNumber: get(935, 944),
      general: get(3882, 3911),
      extraText1: get(3912, 3961),
      extraText2: get(3962, 4011),
      invoices
    };
  };

  private handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
  
    try {
      const text = await file.text();
      const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
      const parsedLines = lines.map(this.parseLine);
  
      // נניח שהפענוח תקין אם יש שורות ויש שם כלשהו בשורה הראשונה
      const isValid = parsedLines.length > 0 && parsedLines[0].name.length > 0;
  
      debugger;
      if (isValid) {
        this.setState({ lines: parsedLines, fileUploaded: true, errorMessage: null });
      } else {
        this.setState({ lines: [], fileUploaded: false, errorMessage: 'הקובץ שהועלה אינו תקין או שלא עבר המרה כהלכה.' });
      }
    } catch (err) {
      this.setState({ lines: [], fileUploaded: false, errorMessage: 'אירעה שגיאה בעת קריאת הקובץ.' });
    }
  };
  
  public render(): React.ReactElement {
    const { lines, fileUploaded, errorMessage } = this.state;

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
          <label
            htmlFor="upload"
            style={{
              display: 'inline-block',
              padding: '0.5em 1em',
              backgroundColor: '#0078d4',
              color: '#fff',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '14px'
            }}
          >
            העלה קובץ
          </label>
          <input
            id="upload"
            type="file"
            onChange={this.handleFile}
            style={{ display: 'none' }}
          />
        </div>

        {lines.length > 0 && (
          <>
            <div style={{ display: 'flex', justifyContent: 'flex-start', marginBottom: '1em' }}>
              <button
                onClick={() => alert('שמירת נתונים תתבצע כאן')}
                style={{
                  padding: '0.5em 1em',
                  backgroundColor: '#107c10',
                  color: '#fff',
                  border: 'none',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  fontSize: '14px'
                }}
              >
                עדכן נתונים
              </button>
            </div>

            <table style={{
              width: '100%',
              borderCollapse: 'collapse',
              fontSize: '13px',
              direction: 'rtl'
            }}>
              <thead style={{ backgroundColor: 'darkseagreen' }}>
                <tr>
                  {[
                    'שם', 'כתובת', 'עיר', 'טלפון', 'מס\' זיהוי', 'מס\' חשבון',
                    'תאריך יצירה', 'מס\' הודעת זיכוי', 'תאריך הודעה',
                    'חשבוניות', 'תיאור', 'לבירורים', 'מייל'
                  ].map((header, i) => (
                    <th key={i} style={{ border: '1px solid #ccc', padding: '8px', textAlign: 'right' }}>{header}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {lines.map((line, index) => (
                  <tr key={index} style={{ backgroundColor: index % 2 === 0 ? '#fff' : '#f9f9f9' }}>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.name}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.address}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.city}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.telephone}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.idNumber}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.accountNumber}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.creationDate}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.creditNoteNumber}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.creditNoteDate}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>
                      <ul style={{ paddingInlineStart: '1.2em', margin: 0 }}>
                        {line.invoices.map((inv, i) => (
                          <li key={i} style={{ marginBottom: '4px' }}>{inv}</li>
                        ))}
                      </ul>
                    </td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.general}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.extraText1}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px' }}>{line.extraText2}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </>
        )}
      </div>
    );
  }
}
