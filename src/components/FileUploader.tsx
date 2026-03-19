import { useState } from 'react';
import type { ParsedFile } from '../utils/excelUtils';
import { parseExcelFile } from '../utils/excelUtils';
import styles from './FileUploader.module.css';

interface FileUploaderProps {
  onFilesAdded: (files: ParsedFile[]) => void;
  anonymize: boolean;
  onAnonymizeChange: (value: boolean) => void;
}

export const FileUploader = ({ onFilesAdded, anonymize, onAnonymizeChange }: FileUploaderProps) => {
  const [isLoading, setIsLoading] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const isExcelFile = (file: File): boolean => {
    return (
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel' ||
      file.name.endsWith('.xlsx') ||
      file.name.endsWith('.xls')
    );
  };

  const processFiles = async (fileList: FileList) => {
    setError(null);
    setSuccessMessage(null);

    if (fileList.length === 0) {
      setError('Ingen filer valgt. Vennligst velg Excel-filer.');
      return;
    }

    const invalidFiles: string[] = [];
    const validFiles: File[] = [];

    // Validate files first
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      if (!isExcelFile(file)) {
        invalidFiles.push(`"${file.name}" er ikke en gyldig Excel-fil (.xlsx eller .xls)`);
      } else {
        validFiles.push(file);
      }
    }

    if (invalidFiles.length > 0) {
      setError(`Ugyldige filer funnet:\n\n${invalidFiles.join('\n')}\n\nForventet: Excel-filer (.xlsx eller .xls)`);
      return;
    }

    if (validFiles.length === 0) {
      setError('Ingen gyldige Excel-filer funnet. Vennligst velg .xlsx eller .xls filer.');
      return;
    }

    setIsLoading(true);
    try {
      const parsedFiles: ParsedFile[] = [];
      const failedFiles: string[] = [];

      for (let i = 0; i < validFiles.length; i++) {
        const file = validFiles[i];
        try {
          const parsed = await parseExcelFile(file);
          parsedFiles.push(parsed);
        } catch (err) {
          failedFiles.push(`"${file.name}": ${err instanceof Error ? err.message : 'Ukjent feil'}`);
        }
      }

      if (parsedFiles.length > 0) {
        onFilesAdded(parsedFiles);
        setSuccessMessage(`✓ Vellykket lastet ${parsedFiles.length} fil(er)`);
        
        if (failedFiles.length > 0) {
          setError(`Klarte ikke å behandle:\n\n${failedFiles.join('\n')}`);
        }
      } else if (failedFiles.length > 0) {
        setError(`Klarte ikke å behandle alle filer:\n\n${failedFiles.join('\n')}`);
      }
    } catch (err) {
      setError(`Uventet feil: ${err instanceof Error ? err.message : 'Ukjent feil'}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.currentTarget.files;
    if (!files) return;
    await processFiles(files);
    event.currentTarget.value = '';
  };

  const handleDragOver = (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const files = e.dataTransfer.files;
    if (!files) {
      setError('Ingen filer ble sluppet. Vennligst prøv igjen.');
      return;
    }

    await processFiles(files);
  };

  return (
    <div className={styles.uploader}>
      <label
        className={`${styles.dropZone} ${isDragging ? styles.dragging : ''} ${isLoading ? styles.loading : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <input
          type="file"
          multiple
          accept=".xlsx,.xls"
          onChange={handleFileSelect}
          disabled={isLoading}
          className={styles.input}
        />
        <div className={styles.uploadIcon}>
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="17 8 12 3 7 8"></polyline>
            <line x1="12" y1="3" x2="12" y2="15"></line>
          </svg>
        </div>
        <p className={styles.mainText}>
          {isLoading ? 'Behandler filer...' : 'Klikk for å velge eller dra Excel-filer hit'}
        </p>
        <p className={styles.subText}>
          Støtter .xlsx og .xls filer
        </p>
      </label>

      <label className={styles.anonymizeLabel}>
        <input
          type="checkbox"
          checked={anonymize}
          onChange={(e) => onAnonymizeChange(e.target.checked)}
          className={styles.anonymizeCheckbox}
        />
        <span>Anonymiser elevnavn</span>
      </label>

      {error && (
        <div className={styles.errorMessage}>
          <strong>❌ Error:</strong>
          <pre>{error}</pre>
        </div>
      )}

      {successMessage && (
        <div className={styles.successMessage}>
          {successMessage}
        </div>
      )}
    </div>
  );
};
