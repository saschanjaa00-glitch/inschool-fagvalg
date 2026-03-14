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
  'matematikk2p',
  'matematikks1',
  'matematikkr1',
  'fremmedsprak',
  'reserve',
];
const FIELD_LABELS: Record<string, string> = {
  navn: 'Navn',
  klasse: 'Klasse',
  blokkmatvg2: 'Mattevalg-kolonne (R1/S1/2P)',
  matematikk2p: 'Matematikk 2P',
  matematikks1: 'Matematikk S1',
  matematikkr1: 'Matematikk R1',
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
  // Create reverse mapping: standardField -> fileColumn
  const getReverseMapping = (fileMapping: ColumnMapping): Record<string, string | null> => {
    const reverse: Record<string, string | null> = {};
    Object.entries(fileMapping).forEach(([fileColumn, standardField]) => {
      if (standardField) {
        reverse[standardField] = fileColumn;
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
    if (fileColumn) {
      // Remove this field from any other column first
      Object.keys(mapping).forEach((col) => {
        if (mapping[col] === field) {
          mapping[col] = null;
        }
      });
      // Set the new mapping
      mapping[fileColumn] = field;
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
      
      {files.map((file) => {
        const fileMapping = currentMappings.get(file.id) || {};
        const reverseMapping = getReverseMapping(fileMapping);
        
        const hasCombinedMathColumn = !!reverseMapping.blokkmatvg2;
        const fieldsForFile = STANDARD_FIELDS.filter((field) => {
          if (!hasCombinedMathColumn) {
            return true;
          }

          return field !== 'matematikk2p' && field !== 'matematikks1' && field !== 'matematikkr1';
        });

        return (
          <div key={file.id} className={styles.fileSection}>
            <h3>{file.filename}</h3>
            <div className={styles.mappingGrid}>
              {fieldsForFile.map((field) => (
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
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
};
