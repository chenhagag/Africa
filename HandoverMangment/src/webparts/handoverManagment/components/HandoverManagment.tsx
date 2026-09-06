import * as React from 'react';
import styles from './HandoverManagment.module.scss';
import type { IHandoverManagmentProps, IApartment } from './IHandoverManagmentProps';
import { BuildingService } from '../services/BuildingService';
import { ComboBox, IComboBoxOption, IComboBox } from '@fluentui/react/lib/ComboBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

const SALE_STATUSES = ['פנויה', 'מכורה', 'בעלים', 'הולד'];
const FINISH_STATUSES_LEGEND = ['גמר מינימלי', 'גמר מלא סטנדרט', 'גמר מלא טיוב תכנון', 'דירת שיווק'];

type TabKey = 'towers' | 'summary' | 'handoverStatus' | 'typeBreakdown';

const HANDOVER_STATUSES = [
  'זומן לטרום',
  'טרום מסירה',
  'טרום מסירה מהנדס',
  'מסירת חזקה',
  'מסירת חזקה מהנדס'
];

// Room groups order for type breakdown
const ROOM_ORDER = [2, 3, 4, 5, 6];

interface IHandoverManagmentState {
  apartments: IApartment[];
  lists: IComboBoxOption[];
  selectedList: string;
  loading: boolean;
  error: string;
  activeTab: TabKey;
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
      activeTab: 'towers'
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
    this._service.getApartments(selectedList)
      .then(apartments => {
        apartments.sort((a, b) => a.apartmentNumber - b.apartmentNumber);
        this.setState({ apartments, loading: false });
      })
      .catch(err => {
        this.setState({ error: `שגיאה בטעינת דירות: ${err.message}`, loading: false });
      });
  }

  private _getApartmentColor(apt: IApartment): string {
    return this.props.saleColors[apt.saleStatus] || this.props.saleColors['_default'] || '#BDBDBD';
  }

  private _getFinishOutline(apt: IApartment): string | undefined {
    const color = this.props.finishColors[apt.finishStatus];
    return color ? `3px solid ${color}` : undefined;
  }

  private _buildTooltip(apt: IApartment): string {
    const lines: string[] = [`דירה ${apt.apartmentNumber}`];
    if (apt.apartmentType) lines.push(`טיפוס: ${apt.apartmentType}`);
    if (apt.rooms) lines.push(`חדרים: ${apt.rooms}`);
    if (apt.standard) lines.push(`סוג: ${apt.standard}`);
    if (apt.saleStatus) lines.push(`מכירה: ${apt.saleStatus}`);
    if (apt.handoverStatus) lines.push(`מסירה: ${apt.handoverStatus}`);
    if (apt.finishStatus) lines.push(`גמר: ${apt.finishStatus}`);
    return lines.join('\n');
  }

  private _onListChange = (_event: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option) {
      this.setState({ selectedList: option.key as string }, () => this._loadApartments());
    }
  }

  private _onApartmentClick(apartment: IApartment): void {
    const url = this._service.getItemFormUrl(this.state.selectedList, apartment.id);
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

  private _groupByTower(apartments: IApartment[]): { south: IApartment[]; north: IApartment[]; middle: IApartment[]; unknown: IApartment[] } {
    const south: IApartment[] = [];
    const north: IApartment[] = [];
    const middle: IApartment[] = [];
    const unknown: IApartment[] = [];

    apartments.forEach(apt => {
      if (apt.tower === 'דרומי') {
        south.push(apt);
      } else if (apt.tower === 'צפוני') {
        north.push(apt);
      } else if (apt.tower === 'אמצע') {
        middle.push(apt);
      } else {
        unknown.push(apt);
      }
    });

    return { south, north, middle, unknown };
  }

  private _buildSummary(apartments: IApartment[]): {
    types: string[];
    south: Record<string, { available: number; sold: number; owner: number; hold: number; total: number }>;
    north: Record<string, { available: number; sold: number; owner: number; hold: number; total: number }>;
    totalSouth: { available: number; sold: number; owner: number; hold: number; total: number };
    totalNorth: { available: number; sold: number; owner: number; hold: number; total: number };
    totalProject: number;
  } {
    const emptyCount = (): { available: number; sold: number; owner: number; hold: number; total: number } =>
      ({ available: 0, sold: 0, owner: 0, hold: 0, total: 0 });

    const south: Record<string, { available: number; sold: number; owner: number; hold: number; total: number }> = {};
    const north: Record<string, { available: number; sold: number; owner: number; hold: number; total: number }> = {};
    const typesSet: Record<string, boolean> = {};

    apartments.forEach(apt => {
      if (apt.isPublicUnit) return;
      const type = apt.apartmentType || '—';
      typesSet[type] = true;
      const isSouth = apt.tower === 'דרומי';
      const isNorth = apt.tower === 'צפוני';
      const bucket = isSouth ? south : isNorth ? north : null;
      if (!bucket) return;

      if (!bucket[type]) bucket[type] = emptyCount();
      bucket[type].total++;

      const status = apt.saleStatus;
      if (status === 'פנויה') bucket[type].available++;
      else if (status === 'מכורה') bucket[type].sold++;
      else if (status === 'בעלים') bucket[type].owner++;
      else if (status === 'הולד') bucket[type].hold++;
    });

    // Sort types: by room prefix then letter
    const types = Object.keys(typesSet);
    types.sort();

    const totalSouth = emptyCount();
    const totalNorth = emptyCount();
    types.forEach(t => {
      if (south[t]) {
        totalSouth.available += south[t].available;
        totalSouth.sold += south[t].sold;
        totalSouth.owner += south[t].owner;
        totalSouth.hold += south[t].hold;
        totalSouth.total += south[t].total;
      }
      if (north[t]) {
        totalNorth.available += north[t].available;
        totalNorth.sold += north[t].sold;
        totalNorth.owner += north[t].owner;
        totalNorth.hold += north[t].hold;
        totalNorth.total += north[t].total;
      }
    });

    return {
      types,
      south,
      north,
      totalSouth,
      totalNorth,
      totalProject: totalSouth.total + totalNorth.total
    };
  }

  private _renderSummaryTable(apartments: IApartment[]): React.ReactElement {
    const s = this._buildSummary(apartments);
    const emptyRow = { available: 0, sold: 0, owner: 0, hold: 0, total: 0 };

    return (
      <div className={styles.summarySection}>
        <h3 className={styles.summaryTitle}>סיכום מלאי דירות</h3>
        <div className={styles.summaryTableWrapper}>
          <table className={styles.summaryTable}>
            <thead>
              <tr>
                <th colSpan={6} className={styles.thSouth}>מגדל דרומי</th>
                <th className={styles.thProject}>כל הפרויקט</th>
                <th colSpan={6} className={styles.thNorth}>מגדל צפוני</th>
              </tr>
              <tr>
                <th className={styles.thType}>טיפוס</th>
                <th className={styles.thAvailable}>פנויות</th>
                <th className={styles.thSold}>מכורות</th>
                <th className={styles.thOwner}>בעלים</th>
                <th className={styles.thHold}>הולד</th>
                <th className={styles.thTotal}>{'סה"כ'}</th>
                <th className={styles.thProject}>{'סה"כ'}</th>
                <th className={styles.thTotal}>{'סה"כ'}</th>
                <th className={styles.thHold}>הולד</th>
                <th className={styles.thOwner}>בעלים</th>
                <th className={styles.thSold}>מכורות</th>
                <th className={styles.thAvailable}>פנויות</th>
                <th className={styles.thType}>טיפוס</th>
              </tr>
            </thead>
            <tbody>
              {s.types.map((type: string) => {
                const sr = s.south[type] || emptyRow;
                const nr = s.north[type] || emptyRow;
                const projectTotal = sr.total + nr.total;
                return (
                  <tr key={type}>
                    <td className={styles.tdType}>{type}</td>
                    <td className={styles.tdAvailable}>{sr.available || ''}</td>
                    <td className={styles.tdSold}>{sr.sold || ''}</td>
                    <td className={styles.tdOwner}>{sr.owner || ''}</td>
                    <td className={styles.tdHold}>{sr.hold || ''}</td>
                    <td className={styles.tdTotal}>{sr.total}</td>
                    <td className={styles.tdProject}>{projectTotal}</td>
                    <td className={styles.tdTotal}>{nr.total}</td>
                    <td className={styles.tdHold}>{nr.hold || ''}</td>
                    <td className={styles.tdOwner}>{nr.owner || ''}</td>
                    <td className={styles.tdSold}>{nr.sold || ''}</td>
                    <td className={styles.tdAvailable}>{nr.available || ''}</td>
                    <td className={styles.tdType}>{type}</td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className={styles.totalRow}>
                <td className={styles.tdType}>{'סה"כ'}</td>
                <td className={styles.tdAvailable}>{s.totalSouth.available}</td>
                <td className={styles.tdSold}>{s.totalSouth.sold}</td>
                <td className={styles.tdOwner}>{s.totalSouth.owner}</td>
                <td className={styles.tdHold}>{s.totalSouth.hold}</td>
                <td className={styles.tdTotal}>{s.totalSouth.total}</td>
                <td className={styles.tdProject}>{s.totalProject}</td>
                <td className={styles.tdTotal}>{s.totalNorth.total}</td>
                <td className={styles.tdHold}>{s.totalNorth.hold}</td>
                <td className={styles.tdOwner}>{s.totalNorth.owner}</td>
                <td className={styles.tdSold}>{s.totalNorth.sold}</td>
                <td className={styles.tdAvailable}>{s.totalNorth.available}</td>
                <td className={styles.tdType}>{'סה"כ'}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  }

  private _renderHandoverStatus(apartments: IApartment[]): React.ReactElement {
    const towers = this._groupByTower(apartments.filter(a => !a.isPublicUnit));

    const countByStatus = (apts: IApartment[]): Record<string, number> => {
      const counts: Record<string, number> = {};
      HANDOVER_STATUSES.forEach(s => { counts[s] = 0; });
      apts.forEach(apt => {
        if (apt.handoverStatus && counts[apt.handoverStatus] !== undefined) {
          counts[apt.handoverStatus]++;
        }
      });
      return counts;
    };

    const southCounts = countByStatus(towers.south);
    const northCounts = countByStatus(towers.north);

    return (
      <div className={styles.summarySection}>
        <h3 className={styles.summaryTitle}>סטטוס מסירות</h3>
        <div className={styles.summaryTableWrapper}>
          <table className={styles.summaryTable}>
            <thead>
              <tr>
                <th colSpan={2} className={styles.thSouth}>מגדל דרומי</th>
                <th className={styles.thProject}>סטטוס</th>
                <th colSpan={2} className={styles.thNorth}>מגדל צפוני</th>
              </tr>
            </thead>
            <tbody>
              {HANDOVER_STATUSES.map((status: string) => (
                <tr key={status}>
                  <td className={styles.tdTotal}>{southCounts[status]}</td>
                  <td className={styles.tdType}>{status}</td>
                  <td className={styles.tdProject}>{southCounts[status] + northCounts[status]}</td>
                  <td className={styles.tdType}>{status}</td>
                  <td className={styles.tdTotal}>{northCounts[status]}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className={styles.totalRow}>
                <td className={styles.tdTotal}>
                  {HANDOVER_STATUSES.reduce((sum, s) => sum + southCounts[s], 0)}
                </td>
                <td className={styles.tdType}>{'סה"כ'}</td>
                <td className={styles.tdProject}>
                  {HANDOVER_STATUSES.reduce((sum, s) => sum + southCounts[s] + northCounts[s], 0)}
                </td>
                <td className={styles.tdType}>{'סה"כ'}</td>
                <td className={styles.tdTotal}>
                  {HANDOVER_STATUSES.reduce((sum, s) => sum + northCounts[s], 0)}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  }

  private _renderTypeBreakdown(apartments: IApartment[]): React.ReactElement {
    const towers = this._groupByTower(apartments.filter(a => !a.isPublicUnit));

    interface ITypeRow {
      label: string;
      southSold: number; southOwner: number; southTotal: number; southAvailable: number; southHold: number;
      northSold: number; northOwner: number; northTotal: number; northAvailable: number; northHold: number;
      grandTotal: number;
      isSubtotal?: boolean;
    }

    const buildRows = (): ITypeRow[] => {
      const rows: ITypeRow[] = [];
      const standards = ['סטנדרט', 'פרימיום', 'פנטהאוס'];
      const standardLabels: Record<string, string> = {
        'סטנדרט': 'סטנדרט',
        'פרימיום': 'פרימיום',
        'פנטהאוס': 'פנטהאוס'
      };

      standards.forEach(std => {
        const subtotal: ITypeRow = {
          label: `סה"כ ${standardLabels[std]}`,
          southSold: 0, southOwner: 0, southTotal: 0, southAvailable: 0, southHold: 0,
          northSold: 0, northOwner: 0, northTotal: 0, northAvailable: 0, northHold: 0,
          grandTotal: 0, isSubtotal: true
        };

        ROOM_ORDER.forEach(rooms => {
          const southApts = towers.south.filter(a => a.rooms === rooms && a.standard === std);
          const northApts = towers.north.filter(a => a.rooms === rooms && a.standard === std);
          if (southApts.length === 0 && northApts.length === 0) return;

          const count = (apts: IApartment[], status: string): number =>
            apts.filter(a => a.saleStatus === status).length;

          const row: ITypeRow = {
            label: `${rooms} חד' ${standardLabels[std]}`,
            southSold: count(southApts, 'מכורה'), southOwner: count(southApts, 'בעלים'),
            southTotal: southApts.length, southAvailable: count(southApts, 'פנויה'),
            southHold: count(southApts, 'הולד'),
            northSold: count(northApts, 'מכורה'), northOwner: count(northApts, 'בעלים'),
            northTotal: northApts.length, northAvailable: count(northApts, 'פנויה'),
            northHold: count(northApts, 'הולד'),
            grandTotal: southApts.length + northApts.length
          };
          rows.push(row);

          subtotal.southSold += row.southSold; subtotal.southOwner += row.southOwner;
          subtotal.southTotal += row.southTotal; subtotal.southAvailable += row.southAvailable;
          subtotal.southHold += row.southHold;
          subtotal.northSold += row.northSold; subtotal.northOwner += row.northOwner;
          subtotal.northTotal += row.northTotal; subtotal.northAvailable += row.northAvailable;
          subtotal.northHold += row.northHold;
          subtotal.grandTotal += row.grandTotal;
        });

        if (subtotal.grandTotal > 0) {
          rows.push(subtotal);
        }
      });

      return rows;
    };

    const rows = buildRows();
    const grandTotal = rows.filter(r => r.isSubtotal).reduce((sum, r) => sum + r.grandTotal, 0);

    return (
      <div className={styles.summarySection}>
        <h3 className={styles.summaryTitle}>פירוט מלאי לפי סוג</h3>
        <div className={styles.summaryTableWrapper}>
          <table className={styles.summaryTable}>
            <thead>
              <tr>
                <th rowSpan={2} className={styles.thType}>טיפוס</th>
                <th colSpan={5} className={styles.thSouth}>מגדל דרומי</th>
                <th rowSpan={2} className={styles.thProject}>{'סה"כ'}</th>
                <th colSpan={5} className={styles.thNorth}>מגדל צפוני</th>
              </tr>
              <tr>
                <th className={styles.thSold}>מכורות</th>
                <th className={styles.thOwner}>בעלים</th>
                <th className={styles.thTotal}>{'סה"כ'}</th>
                <th className={styles.thAvailable}>פנויות</th>
                <th className={styles.thHold}>הולד</th>
                <th className={styles.thSold}>מכורות</th>
                <th className={styles.thOwner}>בעלים</th>
                <th className={styles.thTotal}>{'סה"כ'}</th>
                <th className={styles.thAvailable}>פנויות</th>
                <th className={styles.thHold}>הולד</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row: ITypeRow, idx: number) => (
                <tr key={idx} className={row.isSubtotal ? styles.totalRow : ''}>
                  <td className={styles.tdType}>{row.label}</td>
                  <td className={styles.tdSold}>{row.southSold || ''}</td>
                  <td className={styles.tdOwner}>{row.southOwner || ''}</td>
                  <td className={styles.tdTotal}>{row.southTotal}</td>
                  <td className={styles.tdAvailable}>{row.southAvailable || ''}</td>
                  <td className={styles.tdHold}>{row.southHold || ''}</td>
                  <td className={styles.tdProject}>{row.grandTotal}</td>
                  <td className={styles.tdSold}>{row.northSold || ''}</td>
                  <td className={styles.tdOwner}>{row.northOwner || ''}</td>
                  <td className={styles.tdTotal}>{row.northTotal}</td>
                  <td className={styles.tdAvailable}>{row.northAvailable || ''}</td>
                  <td className={styles.tdHold}>{row.northHold || ''}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className={styles.totalRow}>
                <td className={styles.tdType}>{'סה"כ כללי'}</td>
                <td className={styles.tdSold}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.southSold, 0)}</td>
                <td className={styles.tdOwner}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.southOwner, 0)}</td>
                <td className={styles.tdTotal}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.southTotal, 0)}</td>
                <td className={styles.tdAvailable}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.southAvailable, 0)}</td>
                <td className={styles.tdHold}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.southHold, 0)}</td>
                <td className={styles.tdProject}>{grandTotal}</td>
                <td className={styles.tdSold}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.northSold, 0)}</td>
                <td className={styles.tdOwner}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.northOwner, 0)}</td>
                <td className={styles.tdTotal}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.northTotal, 0)}</td>
                <td className={styles.tdAvailable}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.northAvailable, 0)}</td>
                <td className={styles.tdHold}>{rows.filter(r => r.isSubtotal).reduce((s, r) => s + r.northHold, 0)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    );
  }

  private _getGlobalFloorRange(apartments: IApartment[]): { min: number; max: number } {
    let min = Infinity;
    let max = -Infinity;
    apartments.forEach(apt => {
      if (apt.floor > 0) {
        if (apt.floor < min) min = apt.floor;
        if (apt.floor > max) max = apt.floor;
      }
    });
    return { min: min === Infinity ? 1 : min, max: max === -Infinity ? 1 : max };
  }

  private _renderTower(towerApartments: IApartment[], towerName: string, globalMin: number, globalMax: number): React.ReactElement {
    const floorsMap = this._groupByFloor(towerApartments);
    const floorCount = floorsMap.size;

    // Fill gaps: from globalMax down to globalMin, include missing floors as empty
    const allFloors: number[] = [];
    for (let f = globalMax; f >= globalMin; f--) {
      allFloors.push(f);
    }

    return (
      <div className={styles.towerWrapper}>
        <div className={styles.towerHeader}>
          <Icon iconName="CityNext2" className={styles.towerIcon} />
          <span className={styles.towerName}>{towerName}</span>
          <div className={styles.towerStats}>
            <span className={styles.towerStat}>{towerApartments.filter(a => !a.isPublicUnit).length} דירות</span>
            <span className={styles.towerStat}>{floorCount} קומות</span>
          </div>
        </div>
        <div className={styles.buildingContainer}>
          <div className={styles.roof} />
          {allFloors.map((floorNum: number) => {
            const floorApartments = floorsMap.get(floorNum);
            const isEmpty = !floorApartments || floorApartments.length === 0;

            return (
              <div key={floorNum} className={`${styles.floor}${isEmpty ? ` ${styles.emptyFloor}` : ''}`}>
                <span className={styles.floorLabel}>{floorNum}</span>
                <div className={styles.floorApartments}>
                  {!isEmpty && floorApartments!.map((apt: IApartment) => {
                    if (apt.isPublicUnit) {
                      const puColor = this.props.publicUnitColors[apt.publicUnitType] || '#78909c';
                      return (
                        <div
                          key={apt.id}
                          className={styles.publicUnit}
                          style={{ flex: apt.spaceUnits, backgroundColor: puColor }}
                          onClick={() => this._onApartmentClick(apt)}
                          title={`${apt.title} (${apt.publicUnitType})`}
                        >
                          <span className={styles.publicUnitName}>{apt.title}</span>
                        </div>
                      );
                    }
                    const color = this._getApartmentColor(apt);
                    const outline = this._getFinishOutline(apt);
                    const hasHandover = !!apt.handoverStatus;
                    return (
                      <div
                        key={apt.id}
                        className={`${styles.apartment}${hasHandover ? ` ${styles.hasHandover}` : ''}`}
                        style={{ backgroundColor: color, outline, flex: apt.spaceUnits }}
                        onClick={() => this._onApartmentClick(apt)}
                        title={this._buildTooltip(apt)}
                      >
                        <span className={styles.apartmentNumber}>{apt.apartmentNumber}</span>
                        {apt.apartmentType && (
                          <span className={styles.apartmentType}>{apt.apartmentType}</span>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  private _renderMiddle(middleItems: IApartment[], globalMin: number, globalMax: number): React.ReactElement {
    const floorsMap = this._groupByFloor(middleItems);

    const allFloors: number[] = [];
    for (let f = globalMax; f >= globalMin; f--) {
      allFloors.push(f);
    }

    return (
      <div className={styles.middleColumn}>
        <div className={styles.middleHeader} />
        <div className={styles.middleBody}>
          <div className={styles.middleRoofSpacer} />
          {allFloors.map((floorNum: number) => {
            const items = floorsMap.get(floorNum);
            const hasItems = items && items.length > 0;

            if (!hasItems) {
              return <div key={floorNum} className={styles.middleFloorEmpty} />;
            }

            return (
              <div key={floorNum} className={styles.middleFloor}>
                {items!.map((apt: IApartment) => {
                  const puColor = this.props.publicUnitColors[apt.publicUnitType] || '#78909c';
                  return (
                    <div
                      key={apt.id}
                      className={styles.middleUnit}
                      style={{ backgroundColor: puColor }}
                      title={`${apt.title} (${apt.publicUnitType})`}
                    >
                      <span className={styles.publicUnitName}>{apt.title}</span>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IHandoverManagmentProps> {
    const { apartments, lists, selectedList, loading, error, activeTab } = this.state;
    const towers = this._groupByTower(apartments);
    const hasTowerData = towers.south.length > 0 || towers.north.length > 0;
    const hasMiddle = towers.middle.length > 0;
    const floorRange = this._getGlobalFloorRange(apartments);

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
                  <span className={styles.statValue}>{apartments.length}</span>
                  <span className={styles.statLabel}>דירות</span>
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
            <div className={styles.tabs}>
              <button
                className={`${styles.tab}${activeTab === 'towers' ? ` ${styles.tabActive}` : ''}`}
                onClick={() => this.setState({ activeTab: 'towers' })}
              >
                <Icon iconName="CityNext2" className={styles.tabIcon} />
                מגדלים
              </button>
              <button
                className={`${styles.tab}${activeTab === 'summary' ? ` ${styles.tabActive}` : ''}`}
                onClick={() => this.setState({ activeTab: 'summary' })}
              >
                <Icon iconName="Table" className={styles.tabIcon} />
                סיכום מלאי דירות
              </button>
              <button
                className={`${styles.tab}${activeTab === 'handoverStatus' ? ` ${styles.tabActive}` : ''}`}
                onClick={() => this.setState({ activeTab: 'handoverStatus' })}
              >
                <Icon iconName="TaskManager" className={styles.tabIcon} />
                סטטוס מסירות
              </button>
              <button
                className={`${styles.tab}${activeTab === 'typeBreakdown' ? ` ${styles.tabActive}` : ''}`}
                onClick={() => this.setState({ activeTab: 'typeBreakdown' })}
              >
                <Icon iconName="StackedBarChart" className={styles.tabIcon} />
                פירוט לפי סוג
              </button>
            </div>

            {activeTab === 'towers' && (
              <>
                <div className={styles.legend}>
                  <span className={styles.legendTitle}>סטטוס מכירה:</span>
                  {SALE_STATUSES.map((status: string) => (
                    <div key={status} className={styles.legendItem}>
                      <div className={styles.legendColor} style={{ backgroundColor: this.props.saleColors[status] || '#BDBDBD' }} />
                      <span>{status}</span>
                    </div>
                  ))}
                </div>
                <div className={styles.legend}>
                  <span className={styles.legendTitle}>סטטוס גמר:</span>
                  {FINISH_STATUSES_LEGEND.map((status: string) => (
                    <div key={status} className={styles.legendItem}>
                      <div className={styles.legendColor} style={{ border: `3px solid ${this.props.finishColors[status] || '#999'}`, backgroundColor: 'transparent' }} />
                      <span>{status}</span>
                    </div>
                  ))}
                </div>

                {hasTowerData ? (
                  <div className={styles.towersContainer}>
                    {towers.south.length > 0 && this._renderTower(towers.south, 'מגדל דרומי', floorRange.min, floorRange.max)}
                    {hasMiddle && this._renderMiddle(towers.middle, floorRange.min, floorRange.max)}
                    {towers.north.length > 0 && this._renderTower(towers.north, 'מגדל צפוני', floorRange.min, floorRange.max)}
                  </div>
                ) : (
                  <div className={styles.towersContainer}>
                    {this._renderTower(apartments, 'כל הדירות', floorRange.min, floorRange.max)}
                  </div>
                )}

                {towers.unknown.length > 0 && hasTowerData && (
                  <div className={styles.unknownTower}>
                    <span>{towers.unknown.length} דירות ללא שיוך מגדל</span>
                  </div>
                )}
              </>
            )}

            {activeTab === 'summary' && this._renderSummaryTable(apartments)}

            {activeTab === 'handoverStatus' && this._renderHandoverStatus(apartments)}

            {activeTab === 'typeBreakdown' && this._renderTypeBreakdown(apartments)}
          </>
        )}

        {!loading && !error && apartments.length === 0 && selectedList && (
          <div className={styles.loading}>לא נמצאו דירות ברשימה זו</div>
        )}
      </div>
    );
  }
}
