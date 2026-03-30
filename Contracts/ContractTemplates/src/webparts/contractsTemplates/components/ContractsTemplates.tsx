import * as React from "react";
import styles from "./ContractsTemplates.module.scss";
import { SPFI } from "@pnp/sp";

import { TemplateService } from "../services/TemplateService";
import { ITemplateType, ITemplateLink } from "../models/models";

import {
  Spinner,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  SearchBox
} from "@fluentui/react";

export interface ITemplatePickerProps {
  sp: SPFI;
  templateTypesListTitle: string;
  templateLinksListTitle: string;
}

type ViewState = "types" | "templates";

export default function ContractsTemplates(props: ITemplatePickerProps) {
  const service = React.useMemo(
    () =>
      new TemplateService(
        props.sp,
        props.templateTypesListTitle,
        props.templateLinksListTitle
      ),
    [props.sp, props.templateTypesListTitle, props.templateLinksListTitle]
  );

  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>();

  const [types, setTypes] = React.useState<ITemplateType[]>([]);
  const [allTemplates, setAllTemplates] = React.useState<ITemplateLink[]>([]);

  const [view, setView] = React.useState<ViewState>("types");
  const [selectedType, setSelectedType] = React.useState<ITemplateType | undefined>();
  const [search, setSearch] = React.useState("");

  React.useEffect(() => {
    (async () => {
      try {
        setLoading(true);
        setError(undefined);

        const [t, all] = await Promise.all([
          service.getTypes(),
          service.getAllTemplates()
        ]);

        setTypes(t);
        setAllTemplates(all);
      } catch (e: any) {
        setError(e?.message || "Failed loading data.");
      } finally {
        setLoading(false);
      }
    })();
  }, [service]);

  const onSelectType = (t: ITemplateType) => {
    setSelectedType(t);
    setView("templates");
    setSearch("");
  };

  const scopedTemplates = React.useMemo(() => {
    if (view === "templates" && selectedType) {
      return allTemplates.filter(t => t.TemplateTypeId === selectedType.Id);
    }
    return allTemplates;
  }, [allTemplates, view, selectedType]);

  const filteredTemplates = React.useMemo(() => {
    const s = (search || "").trim().toLowerCase();
    if (!s) return scopedTemplates;

    return scopedTemplates.filter(x => {
      const title = (x.Title || "").toLowerCase();
      const desc = (x.Description || "").toLowerCase();
      return title.indexOf(s) !== -1 || desc.indexOf(s) !== -1;
    });
  }, [scopedTemplates, search]);

  const hasSearch = !!(search || "").trim();


  const buildNewFileName = (templateId: number) => {
    const now = new Date();
    const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`);
    const ts = `${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  
    return `CONT-${ts}-${templateId}.docx`;
  };
  
  const openWebEdit = (serverRelativeUrl: string) => {
    const path = serverRelativeUrl.startsWith("/") ? serverRelativeUrl : `/${serverRelativeUrl}`;
    window.open(`${window.location.origin}${path}?web=1`, "_blank", "noopener,noreferrer");
  };
  
  
  const onCreateDocument = async (tpl: ITemplateLink) => {
    if (!tpl.TemplateServerRelativeUrl) {
      setError("לתבנית הזו חסר נתיב TemplateServerRelativeUrl.");
      return;
    }
  
    try {
      setLoading(true);
      setError(undefined);
  
      const targetFolder = "/ContractsNEW/Shared Documents";
      const newFileName = buildNewFileName(tpl.Id);
  
      const res = await service.createDocFromTemplateBlob({
        templateServerRelativeUrl: tpl.TemplateServerRelativeUrl,
        targetFolderServerRelativeUrl: targetFolder,
        newFileName
      });
  
      openWebEdit(res.newFileServerRelativeUrl);
    } catch (e: any) {
      setError(e?.message || "שגיאה ביצירת מסמך מהתבנית.");
    } finally {
      setLoading(false);
    }
  };
  

  return (
    <div className={styles.templatePickerRoot}>
      <div className={styles.header}>
        <div className={styles.title}>יצירת מסמך מתבנית</div>

        {view === "templates" && selectedType && (
          <div className={styles.subTitle}>
            סוג נבחר: <b>{selectedType.Title}</b>
          </div>
        )}
      </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
        </MessageBar>
      )}

      {loading && (
        <div className={styles.loading}>
          <Spinner label="טוען..." />
        </div>
      )}

      {/* ====== מסך ראשי: סוגים + חיפוש גלובלי ====== */}
      {!loading && view === "types" && (
        <div>
          <div className={`${styles.actionsRow} ${styles.centered}`}>
            <div className={styles.searchBoxWrap}>
              <div className={styles.searchBoxStyled}>
                <SearchBox
                  placeholder="חיפוש תבנית..."
                  value={search}
                  onChange={(_, val) => setSearch(val || "")}
                />
              </div>
            </div>
          </div>

          {hasSearch && (
            <div className={styles.list}>
              {filteredTemplates.map(tpl => (
                <div key={tpl.Id} className={styles.row}>
                  <div className={styles.rowMain}>
                    <div className={styles.rowTitle}>{tpl.Title}</div>
                    <div className={styles.rowDesc}>
                      סוג: <b>{tpl.TemplateTypeTitle || "לא משויך"}</b>
                    </div>
                  </div>

                  <PrimaryButton
                    text="צור מסמך"
                    onClick={() => onCreateDocument(tpl)}
                    disabled={!tpl.TemplateServerRelativeUrl}
                  />
                </div>
              ))}

              {filteredTemplates.length === 0 && (
                <div className={styles.empty}>לא נמצאו תבניות.</div>
              )}
            </div>
          )}

          <div className={styles.grid}>
            {types.map(t => (
              <button
                key={t.Id}
                className={styles.tile}
                onClick={() => onSelectType(t)}
                type="button"
                title={t.Title}
              >
                <div className={styles.tileTitle}>{t.Title}</div>
                <div className={styles.tileMeta}>לחצ/י להצגת תבניות</div>
              </button>
            ))}

            {types.length === 0 && (
              <div className={styles.empty}>
                לא נמצאו סוגי תבניות ברשימה <b>{props.templateTypesListTitle}</b>.
              </div>
            )}
          </div>
        </div>
      )}

      {/* ====== מסך סוג: חיפוש בתוך סוג ====== */}
      {!loading && view === "templates" && (
        <div className={styles.templatesArea}>
          <div className={styles.actionsRow}>
            <DefaultButton
              text="חזרה לסוגים"
              onClick={() => {
                setView("types");
                setSelectedType(undefined);
                setSearch("");
              }}
            />

            <div className={styles.searchBoxWrap}>
              <div className={styles.searchBoxStyled}>
                <SearchBox
                  placeholder="חיפוש בתוך הסוג..."
                  value={search}
                  onChange={(_, val) => setSearch(val || "")}
                />
              </div>
            </div>
          </div>

          <div className={styles.list}>
            {filteredTemplates.map(tpl => (
              <div key={tpl.Id} className={styles.row}>
                <div className={styles.rowMain}>
                  <div className={styles.rowTitle}>{tpl.Title}</div>
                  {tpl.Description && (
                    <div className={styles.rowDesc}>{tpl.Description}</div>
                  )}
                </div>

                <PrimaryButton
                  text="צור מסמך"
                  onClick={() => onCreateDocument(tpl)}
                  disabled={!tpl.TemplateServerRelativeUrl}
                />
              </div>
            ))}

            {filteredTemplates.length === 0 && (
              <div className={styles.empty}>לא נמצאו תבניות לסוג זה.</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
