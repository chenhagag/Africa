import * as React from 'react';
import styles from './HandoverManagment.module.scss';
import type { IHandoverManagmentProps, IApartment } from './IHandoverManagmentProps';
import { BuildingService } from '../services/BuildingService';
import { ComboBox, IComboBoxOption, IComboBox } from '@fluentui/react/lib/ComboBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

// POC colors
const POC_COLORS = {
  green: '#4CAF50',
  yellow: '#FFC107',
  red: '#F44336'
};

interface IHandoverManagmentState {
  apartments: IApartment[];
  lists: IComboBoxOption[];
  selectedList: string;
  loading: boolean;
  error: string;
  totalApartments: number;
}

export default class HandoverManagment extends React.Component<IHandoverManagmentProps, IHandoverManagmentState> {
  private _service: BuildingService;

  constructor(props: IHandoverManagmentProps) {
    super(props);
    this._service = new BuildingService(props.spHttpClient, props.siteUrl);
    this.state = {
      apartments: [],
      lists: [],
      selectedList: props.listName,
      loading: true,
      error: '',
      totalApartments: 0
    };
  }

  public componentDidMount(): void {
    this._loadLists();
  }

  public componentDidUpdate(prevProps: IHandoverManagmentProps): void {
    if (prevProps.siteUrl !== this.props.siteUrl || prevProps.listName !== this.props.listName) {
      this._service = new BuildingService(this.props.spHttpClient, this.props.siteUrl);
      this.setState({ selectedList: this.props.listName }, () => this._loadLists());
    }
  }

  private _loadLists(): void {
    this._service.getLists()
      .then(lists => {
        const options: IComboBoxOption[] = lists.map(l => ({ key: l.Title, text: l.Title }));
        this.setState({ lists: options }, () => this._loadApartments());
      })
      .catch(err => {
        this.setState({ error: `שגיאה בטעינת רשימות: ${err.message}`, loading: false });
      });
  }

  private _loadApartments(): void {
    const { selectedList } = this.state;
    if (!selectedList) {
      this.setState({ loading: false });
      return;
    }

    this.setState({ loading: true, error: '' });
    this._service.getApartments(selectedList, this.props.apartmentsPerFloor)
      .then(apartments => {
        apartments.sort((a, b) => a.apartmentNumber - b.apartmentNumber);
        this.setState({ apartments, loading: false, totalApartments: apartments.length });
      })
      .catch(err => {
        this.setState({ error: `שגיאה בטעינת דירות: ${err.message}`, loading: false });
      });
  }

  // POC: assign colors based on index position
  private _getPocColor(index: number): string {
    if (index % 7 === 3) return POC_COLORS.red;
    if (index % 5 === 4) return POC_COLORS.yellow;
    return POC_COLORS.green;
  }

  private _getPocStatus(index: number): string {
    if (index % 7 === 3) return 'ממתין לתיקונים';
    if (index % 5 === 4) return 'ממתין לחתימה';
    return 'מוכן למסירה';
  }

  private _onListChange = (_event: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option) {
      this.setState({ selectedList: option.key as string }, () => this._loadApartments());
    }
  }

  private _onApartmentClick(apartment: IApartment): void {
    const url = this._service.getItemFormUrl(this.state.selectedList, apartment.id);
    // Open as popup dialog
    const width = 800;
    const height = 600;
    const left = (window.screen.width - width) / 2;
    const top = (window.screen.height - height) / 2;
    window.open(url, 'ApartmentDetails', `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`);
  }

  private _groupByFloor(apartments: IApartment[]): Map<number, IApartment[]> {
    const floors = new Map<number, IApartment[]>();
    apartments.forEach(apt => {
      if (!floors.has(apt.floor)) {
        floors.set(apt.floor, []);
      }
      floors.get(apt.floor)!.push(apt);
    });
    return floors;
  }

  public render(): React.ReactElement<IHandoverManagmentProps> {
    const { apartments, lists, selectedList, loading, error, totalApartments } = this.state;
    const floors = this._groupByFloor(apartments);
    const sortedFloors: [number, IApartment[]][] = Array.from(floors.entries()).sort((a: [number, IApartment[]], b: [number, IApartment[]]) => b[0] - a[0]);
    const floorCount = sortedFloors.length;

    return (
      <div className={styles.handoverManagment}>
        <div className={styles.headerCard}>
          <div className={styles.headerTop}>
            <Icon iconName="CityNext" className={styles.headerIcon} />
            <h2 className={styles.title}>תצוגת בניין - מסירות</h2>
          </div>
          <div className={styles.headerControls}>
            <div className={styles.comboWrapper}>
              <label className={styles.comboLabel}>בחירת בניין</label>
              <ComboBox
                className={styles.combo}
                placeholder="הקלד לחיפוש בניין..."
                options={lists}
                selectedKey={selectedList}
                onChange={this._onListChange}
                autoComplete="on"
                allowFreeform={true}
              />
            </div>
            {selectedList && !loading && (
              <div className={styles.headerStats}>
                <div className={styles.statItem}>
                  <span className={styles.statValue}>{totalApartments}</span>
                  <span className={styles.statLabel}>דירות</span>
                </div>
                <div className={styles.statItem}>
                  <span className={styles.statValue}>{floorCount}</span>
                  <span className={styles.statLabel}>קומות</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {loading && (
          <div className={styles.loading}>
            <Spinner size={SpinnerSize.large} label="טוען נתונים..." />
          </div>
        )}

        {error && <div className={styles.error}>{error}</div>}

        {!loading && !error && apartments.length > 0 && (
          <>
            <div className={styles.buildingContainer}>
              <div className={styles.roof} />
              {sortedFloors.map(([floorNum, floorApartments]: [number, IApartment[]]) => (
                <div key={floorNum} className={styles.floor}>
                  <span className={styles.floorLabel}>קומה {floorNum}</span>
                  <div className={styles.floorApartments}>
                    {floorApartments.map((apt: IApartment, idx: number) => {
                      const globalIndex = apartments.indexOf(apt);
                      const color = this._getPocColor(globalIndex);
                      const status = this._getPocStatus(globalIndex);
                      return (
                        <div
                          key={apt.id}
                          className={styles.apartment}
                          style={{ backgroundColor: color }}
                          onClick={() => this._onApartmentClick(apt)}
                          title={`דירה ${apt.apartmentNumber} - ${status}`}
                        >
                          <span className={styles.apartmentNumber}>{apt.apartmentNumber}</span>
                          <span className={styles.apartmentStatus}>{status}</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>

            <div className={styles.legend}>
              <div className={styles.legendItem}>
                <div className={styles.legendColor} style={{ backgroundColor: POC_COLORS.green }} />
                <span>מוכן למסירה</span>
              </div>
              <div className={styles.legendItem}>
                <div className={styles.legendColor} style={{ backgroundColor: POC_COLORS.yellow }} />
                <span>ממתין לחתימה</span>
              </div>
              <div className={styles.legendItem}>
                <div className={styles.legendColor} style={{ backgroundColor: POC_COLORS.red }} />
                <span>ממתין לתיקונים</span>
              </div>
            </div>
          </>
        )}

        {!loading && !error && apartments.length === 0 && selectedList && (
          <div className={styles.loading}>לא נמצאו דירות ברשימה זו</div>
        )}
      </div>
    );
  }
}
