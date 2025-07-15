import * as React from 'react';
import { sp } from '@pnp/sp/presets/all';
import styles from './SignaturePage.module.scss';
import type { ISignaturePageProps } from './ISignaturePageProps';

interface ISignaturePageState {
  userName: string;
  loanItemId: number | null;
  currentItemTitle: string;
  otherItems: string[];
  isSigned: boolean;
  isLoading: boolean;
  error?: string;
}

export default class SignaturePage extends React.Component<ISignaturePageProps, ISignaturePageState> {
  private canvasRef = React.createRef<HTMLCanvasElement>();
  private isDrawing = false;

  constructor(props: ISignaturePageProps) {
    super(props);
    this.state = {
      userName: '',
      loanItemId: null,
      currentItemTitle: '',
      otherItems: [],
      isSigned: false,
      isLoading: true
    };

    sp.setup({ spfxContext: this.props.context as any });
  }

  public async componentDidMount(): Promise<void> {
    const itemIdStr = new URLSearchParams(window.location.search).get("ItemID");
    const itemId = itemIdStr ? parseInt(itemIdStr) : null;

    if (!itemId) {
      this.setState({ error: "ItemID לא סופק ב־URL", isLoading: false });
      return;
    }

    try {
      const loanItem = await sp.web.lists.getByTitle("השאלות").items
        .getById(itemId)
        .select("CurrOwner/Title", "CurrOwner/ID", "Item/Title")
        .expand("CurrOwner", "Item")();

      const userId = loanItem.CurrOwner?.ID;
      const userName = loanItem.CurrOwner?.Title;
      const currentItemTitle = loanItem.Item?.Title || "(לא ידוע)";

      const oldLoanItems = await sp.web.lists.getByTitle("השאלות").items
        .filter(`CurrOwnerId eq ${userId} and Status eq 'השאלה פעילה'`)
        .select("ID", "Item/Title")
        .expand("Item")
        .top(50)
        .get();

      const otherItems = oldLoanItems
        .filter(i => i.ID !== itemId)
        .map(i => i.Item?.Title || "(ללא שם)");

      this.setState({
        loanItemId: itemId,
        userName,
        currentItemTitle,
        otherItems,
        isLoading: false
      });

    } catch (error) {
      this.setState({ error: "שגיאה בטעינת נתוני ההשאלה", isLoading: false });
      console.error(error);
    }
  }

  private handleMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const ctx = this.canvasRef.current?.getContext('2d');
    if (!ctx) return;
    ctx.beginPath();
    ctx.moveTo(e.nativeEvent.offsetX, e.nativeEvent.offsetY);
    this.isDrawing = true;
  };

  private handleMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!this.isDrawing) return;
    const ctx = this.canvasRef.current?.getContext('2d');
    if (!ctx) return;
    ctx.lineTo(e.nativeEvent.offsetX, e.nativeEvent.offsetY);
    ctx.stroke();
  };

  private handleMouseUp = () => {
    this.isDrawing = false;
  };

  private clearCanvas = () => {
    const canvas = this.canvasRef.current;
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }
    }
  };

  private dataURLtoBlob = (dataUrl: string): Blob => {
    const arr = dataUrl.split(',');
    const mime = arr[0].match(/:(.*?);/)?.[1];
    const bstr = atob(arr[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) {
      u8arr[n] = bstr.charCodeAt(n);
    }
    return new Blob([u8arr], { type: mime });
  };

  private handleSign = async () => {
    const { loanItemId } = this.state;
    const canvas = this.canvasRef.current;
    if (!loanItemId || !canvas) return;

    try {
      const dataUrl = canvas.toDataURL('image/png');
      const blob = this.dataURLtoBlob(dataUrl);
      const fileName = `signature-${loanItemId}-${Date.now()}.png`;

      const currentUser = await sp.web.currentUser.get();
      const now = new Date().toLocaleDateString('he-IL');
      const signatureText = `נחתם על ידי ${currentUser.Title} בתאריך ${now}`;

      await sp.web.lists
        .getByTitle("השאלות")
        .items.getById(loanItemId)
        .attachmentFiles.add(fileName, blob);

      await sp.web.lists.getByTitle("השאלות").items.getById(loanItemId).update({
        Signature0: signatureText
      });

      this.setState({ isSigned: true });

    } catch (error) {
      this.setState({ error: "שגיאה בשמירת החתימה" });
      console.error(error);
    }
  };

  public render(): React.ReactElement<ISignaturePageProps> {
    const { isLoading, userName, currentItemTitle, otherItems, isSigned, error } = this.state;

    if (isLoading) return <div>טוען נתונים...</div>;
    if (error) return <div className={styles.error}>{error}</div>;

    return (
      <div className={styles.signaturePage}>
        <h2>אישור השאלת ציוד</h2>
        <p><strong>משתמש:</strong> {userName}</p>
        <p><strong>פריט חדש:</strong> <span className={styles.highlight}>{currentItemTitle}</span></p>

        {otherItems.length > 0 && (
          <>
            <h4>פריטים נוספים שהושאלו למשתמש:</h4>
            <ul>
              {otherItems.map((title, idx) => (
                <li key={idx}>{title}</li>
              ))}
            </ul>
          </>
        )}

        {!isSigned ? (
          <div className={styles.signatureBox}>
            <p><strong>חתימה ידנית:</strong></p>
            <canvas
              ref={this.canvasRef}
              width={300}
              height={150}
              className={styles.signatureCanvas}
              onMouseDown={this.handleMouseDown}
              onMouseMove={this.handleMouseMove}
              onMouseUp={this.handleMouseUp}
              onMouseLeave={this.handleMouseUp}
            />
            <div className={styles.buttonGroup}>
              <button className={styles.clearBtn} onClick={this.clearCanvas}>נקה חתימה</button>
              <button className={styles.signButton} onClick={this.handleSign}>חתום על הפריט</button>
            </div>
          </div>
        ) : (
          <p className={styles.success}>✔️ החתימה נקלטה בהצלחה</p>
        )}
      </div>
    );
  }
}
