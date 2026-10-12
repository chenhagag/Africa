import * as React from 'react';
import styles from './HandoverManagment.module.scss';
import { BuildingService, IFieldInfo } from '../services/BuildingService';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import type { IApartment } from './IHandoverManagmentProps';

type ViewKey = 'general' | 'sales' | 'handover' | 'suppliers';

interface IViewConfig {
  key: ViewKey;
  label: string;
  icon: string;
  defaultFields: string[];
}

const VIEW_CONFIGS: IViewConfig[] = [
  {
    key: 'general',
    label: 'כללי',
    icon: 'Info',
    defaultFields: ['msrAppNum', 'Title', 'Floor', 'Tower', 'ApartmentType', 'Rooms', 'Standard', 'ShadowNumber']
  },
  {
    key: 'sales',
    label: 'מכירה ורוכש',
    icon: 'Money',
    defaultFields: ['msrAppStatus', 'BuyerName', 'BuyerPhone', 'BuyerEmail', 'SaleDate', 'ContractNumber']
  },
  {
    key: 'handover',
    label: 'מסירה וגמר',
    icon: 'TaskManager',
    defaultFields: ['HandoverStatus', 'FinishStatus', 'HandoverDate', 'PreHandoverDate', 'EngineerDate', 'Remarks']
  },
  {
    key: 'suppliers',
    label: 'ספקים וקבלנים',
    icon: 'People',
    defaultFields: ['Contractor', 'ContractorPhone', 'Electrician', 'Plumber', 'ACTechnician', 'SupplierRemarks']
  }
];

// Fields to exclude from "add columns" picker (internal SP fields)
const EXCLUDED_FIELDS = new Set([
  'ContentType', 'ContentTypeId', '_ModerationComments', '_ModerationStatus',
  'FileSystemObjectType', 'ServerRedirectedEmbedUri', 'ServerRedirectedEmbedUrl',
  'ComplianceAssetId', 'OData__UIVersionString', 'GUID', 'Attachments',
  'AuthorId', 'EditorId', 'OData__ColorTag', 'AppAuthorId', 'AppEditorId',
  'ID', 'Id'
]);

interface IApartmentDetailsModalProps {
  apartment: IApartment;
  service: BuildingService;
  listName: string;
  onDismiss: () => void;
}

interface IApartmentDetailsModalState {
  activeView: ViewKey;
  itemData: Record<string, unknown> | null;
  listFields: IFieldInfo[];
  loading: boolean;
  error: string;
  viewFields: Record<ViewKey, string[]>;
  showFieldPicker: boolean;
}

export default class ApartmentDetailsModal extends React.Component<IApartmentDetailsModalProps, IApartmentDetailsModalState> {

  constructor(props: IApartmentDetailsModalProps) {
    super(props);

    const viewFields: Record<ViewKey, string[]> = {} as Record<ViewKey, string[]>;
    VIEW_CONFIGS.forEach(v => {
      viewFields[v.key] = [...v.defaultFields];
    });

    this.state = {
      activeView: 'general',
      itemData: null,
      listFields: [],
      loading: true,
      error: '',
      viewFields,
      showFieldPicker: false
    };
  }

  public componentDidMount(): void {
    const { service, listName, apartment } = this.props;

    Promise.all([
      service.getItemAllFields(listName, apartment.id),
      service.getListFields(listName)
    ])
      .then(([itemData, listFields]) => {
        this.setState({ itemData, listFields, loading: false });
      })
      .catch(err => {
        this.setState({ error: `שגיאה בטעינת נתונים: ${err.message}`, loading: false });
      });
  }

  private _getFieldDisplayName(internalName: string): string {
    const field = this.state.listFields.find(f => f.InternalName === internalName);
    return field ? field.Title : internalName;
  }

  private _formatFieldValue(value: unknown): string {
    if (value === null || value === undefined || value === '') return '—';
    if (typeof value === 'boolean') return value ? 'כן' : 'לא';
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  }

  private _getAvailableFields(): IFieldInfo[] {
    const { listFields, viewFields, activeView } = this.state;
    const currentFields = new Set(viewFields[activeView]);
    return listFields.filter(f =>
      !currentFields.has(f.InternalName) &&
      !EXCLUDED_FIELDS.has(f.InternalName)
    );
  }

  private _addField(internalName: string): void {
    const { viewFields, activeView } = this.state;
    const updated = { ...viewFields };
    updated[activeView] = [...updated[activeView], internalName];
    this.setState({ viewFields: updated, showFieldPicker: false });
  }

  private _removeField(internalName: string): void {
    const { viewFields, activeView } = this.state;
    const updated = { ...viewFields };
    updated[activeView] = updated[activeView].filter(f => f !== internalName);
    this.setState({ viewFields: updated });
  }

  private _onEditClick = (): void => {
    const { service, listName, apartment } = this.props;
    const url = service.getEditFormUrl(listName, apartment.id);
    const width = 800;
    const height = 600;
    const left = (window.screen.width - width) / 2;
    const top = (window.screen.height - height) / 2;
    window.open(url, 'ApartmentEdit', `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes`);
  }

  private _onOverlayClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) {
      this.props.onDismiss();
    }
  }

  public render(): React.ReactElement {
    const { apartment, onDismiss } = this.props;
    const { activeView, itemData, loading, error, viewFields, showFieldPicker } = this.state;
    const currentFields = viewFields[activeView];
    const availableFields = this._getAvailableFields();

    return (
      <div className={styles.modalOverlay} onClick={this._onOverlayClick}>
        <div className={styles.modalContainer}>
          {/* Header */}
          <div className={styles.modalHeader}>
            <div className={styles.modalHeaderInfo}>
              <Icon iconName="Home" className={styles.modalHeaderIcon} />
              <h2 className={styles.modalTitle}>דירה {apartment.apartmentNumber}</h2>
              <span className={styles.modalSubtitle}>
                {apartment.tower && `${apartment.tower} · `}קומה {apartment.floor}
                {apartment.apartmentType && ` · ${apartment.apartmentType}`}
              </span>
            </div>
            <div className={styles.modalHeaderActions}>
              <button className={styles.modalEditBtn} onClick={this._onEditClick} title="עריכת פריט">
                <Icon iconName="Edit" />
                <span>עריכה</span>
              </button>
              <button className={styles.modalCloseBtn} onClick={onDismiss} title="סגירה">
                <Icon iconName="Cancel" />
              </button>
            </div>
          </div>

          {/* View Tabs */}
          <div className={styles.modalTabs}>
            {VIEW_CONFIGS.map(view => (
              <button
                key={view.key}
                className={`${styles.modalTab}${activeView === view.key ? ` ${styles.modalTabActive}` : ''}`}
                onClick={() => this.setState({ activeView: view.key, showFieldPicker: false })}
              >
                <Icon iconName={view.icon} className={styles.modalTabIcon} />
                {view.label}
              </button>
            ))}
          </div>

          {/* Content */}
          <div className={styles.modalBody}>
            {loading && (
              <div className={styles.modalLoading}>
                <Spinner size={SpinnerSize.large} label="טוען פרטי דירה..." />
              </div>
            )}

            {error && <div className={styles.modalError}>{error}</div>}

            {!loading && !error && itemData && (
              <>
                <table className={styles.modalFieldsTable}>
                  <tbody>
                    {currentFields.map(fieldName => (
                      <tr key={fieldName} className={styles.modalFieldRow}>
                        <td className={styles.modalFieldLabel}>
                          {this._getFieldDisplayName(fieldName)}
                        </td>
                        <td className={styles.modalFieldValue}>
                          {this._formatFieldValue(itemData[fieldName])}
                        </td>
                        <td className={styles.modalFieldRemove}>
                          <button
                            className={styles.modalRemoveFieldBtn}
                            onClick={() => this._removeField(fieldName)}
                            title="הסר עמודה"
                          >
                            <Icon iconName="Cancel" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* Add columns button */}
                <div className={styles.modalAddFieldSection}>
                  <button
                    className={styles.modalAddFieldBtn}
                    onClick={() => this.setState({ showFieldPicker: !showFieldPicker })}
                  >
                    <Icon iconName="Add" />
                    <span>הוספת עמודות</span>
                  </button>

                  {showFieldPicker && (
                    <div className={styles.fieldPickerDropdown}>
                      <div className={styles.fieldPickerTitle}>בחר עמודה להוספה:</div>
                      <div className={styles.fieldPickerList}>
                        {availableFields.length === 0 && (
                          <div className={styles.fieldPickerEmpty}>אין עמודות נוספות זמינות</div>
                        )}
                        {availableFields.map(field => (
                          <button
                            key={field.InternalName}
                            className={styles.fieldPickerItem}
                            onClick={() => this._addField(field.InternalName)}
                          >
                            <Icon iconName="Add" className={styles.fieldPickerItemIcon} />
                            <span>{field.Title}</span>
                            <span className={styles.fieldPickerItemType}>({field.TypeAsString})</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    );
  }
}
