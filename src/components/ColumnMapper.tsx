import type { ParsedFile, ColumnMapping } from '../utils/excelUtils';
import { getBlokkFields } from '../utils/excelUtils';
import styles from './ColumnMapper.module.css';

interface ColumnMapperProps {
  files: ParsedFile[];
  onMappingChange: (fileId: string, mapping: ColumnMapping) => void;
  currentMappings: Map<string, ColumnMapping>;
  blokkCount: number;
  onBlokkCountChange: (count: number) => void;
}

const STANDARD_FIELDS_BASE = [
  'navn',
  'klasse',
  'blokkmatvg2',
  'fremmedsprak',
  'reserve',
];
const FIELD_LABELS: Record<string, string> = {
  navn: 'Navn',
  klasse: 'Klasse',
  blokkmatvg2: 'Mattevalg (R1/S1/2P)',
  fremmedsprak: 'Fremmedspråk',
  reserve: 'Reserve',
  blokk1: 'Blokk 1',
  blokk2: 'Blokk 2',
  blokk3: 'Blokk 3',
  blokk4: 'Blokk 4',
  blokk5: 'Blokk 5',
  blokk6: 'Blokk 6',
  blokk7: 'Blokk 7',
  blokk8: 'Blokk 8',
};

export const ColumnMapper = ({
  files,
  onMappingChange,
  currentMappings,
  blokkCount,
  onBlokkCountChange,
}: ColumnMapperProps) => {
  const STANDARD_FIELDS = [...STANDARD_FIELDS_BASE, ...getBlokkFields(blokkCount)];
  // Prefix used for blokkmatvg2 mapping keys to allow a column to be mapped
  // to both a blokk field and blokkmatvg2 simultaneously.
  const MATTE_PREFIX = '__mattevalg__';

  // Create reverse mapping: standardField -> fileColumn
  const getReverseMapping = (fileMapping: ColumnMapping): Record<string, string | null> => {
    const reverse: Record<string, string | null> = {};
    Object.entries(fileMapping).forEach(([fileColumn, standardField]) => {
      if (standardField) {
        reverse[standardField] = fileColumn.startsWith(MATTE_PREFIX)
          ? fileColumn.slice(MATTE_PREFIX.length)
          : fileColumn;
      }
    });
    return reverse;
  };

  const handleMappingChange = (fileId: string, field: string, fileColumn: string | null) => {
    const file = files.find((f) => f.id === fileId);
    if (!file) return;

    const mapping: ColumnMapping = {};
    
    // Initialize mapping - set all columns to null first
    file.columns.forEach((col) => {
      mapping[col] = null;
    });
    
    // Copy existing mappings for other fields
    const currentMapping = currentMappings.get(fileId) || {};
    Object.entries(currentMapping).forEach(([col, field]) => {
      if (field && field !== null) {
        mapping[col] = field;
      }
    });
    
    // Update the mapping for this field
    // First, remove this field from any column it was previously assigned to
    Object.keys(mapping).forEach((col) => {
      if (mapping[col] === field) {
        mapping[col] = null;
      }
    });

    if (fileColumn) {
      if (field === 'blokkmatvg2') {
        // Use synthetic key so the blokk mapping on the same column is preserved
        mapping[MATTE_PREFIX + fileColumn] = field;
      } else {
        mapping[fileColumn] = field;
      }
    }
    
    onMappingChange(fileId, mapping);
  };

  return (
    <div className={styles.mapper}>
      <div className={styles.blokkCountSelector}>
        <label htmlFor="blokk-count">Antall Blokk-kolonner:</label>
        <input
          id="blokk-count"
          type="number"
          min="1"
          max="8"
          value={blokkCount}
          onChange={(e) => onBlokkCountChange(Math.max(1, Math.min(8, parseInt(e.target.value) || 4)))}
          className={styles.blokkCountInput}
        />
      </div>
      <p>Kolonner har blitt automatisk oppdaget. Du kan justere tilordningene om nødvendig:</p>
      
      {[...files].sort((a, b) => a.filename.localeCompare(b.filename, 'nb', { numeric: true })).map((file) => {
        const fileMapping = currentMappings.get(file.id) || {};
        const reverseMapping = getReverseMapping(fileMapping);
        
        const fieldsForFile = STANDARD_FIELDS;

        const fieldGroups: Array<{ label: string; fields: string[] }> = [
          { label: 'Elev', fields: fieldsForFile.filter((f) => f === 'navn' || f === 'klasse') },
          { label: 'Blokker', fields: fieldsForFile.filter((f) => f.startsWith('blokk') && f !== 'blokkmatvg2') },
          { label: 'Matematikk', fields: fieldsForFile.filter((f) => f === 'blokkmatvg2') },
          { label: 'Annet', fields: fieldsForFile.filter((f) => f === 'fremmedsprak' || f === 'reserve') },
        ].filter((g) => g.fields.length > 0);

        const renderField = (field: string) => (
          <div key={field} className={styles.mappingRow}>
            <label>{FIELD_LABELS[field]}</label>
            <select
              value={reverseMapping[field] || ''}
              onChange={(e) => handleMappingChange(file.id, field, e.target.value || null)}
              className={styles.select}
            >
              <option value="">-- Ikke tilordnet --</option>
              {file.columns.map((col, index) => (
                <option key={`${col}-${index}`} value={col}>
                  {col}
                </option>
              ))}
            </select>
          </div>
        );

        return (
          <div key={file.id} className={styles.fileSection}>
            <h3>{file.filename}</h3>
            <div className={styles.groupsRow}>
              {fieldGroups.map((group) => (
                <div key={group.label} className={styles.fieldGroup}>
                  <span className={styles.fieldGroupLabel}>{group.label}</span>
                  <div className={styles.mappingGrid}>
                    {group.fields.map((field, idx) => (
                      <>{idx === 4 && <div className={styles.mappingGridBreak} />}{renderField(field)}</>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
};
