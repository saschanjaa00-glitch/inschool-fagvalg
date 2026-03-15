import { mapSubjectToCode } from './subjectCodeMapping';

let xlsxModulePromise: Promise<typeof import('xlsx')> | null = null;

export const loadXlsx = async () => {
  if (!xlsxModulePromise) {
    xlsxModulePromise = import('xlsx');
  }

  return xlsxModulePromise;
};

export interface ParsedFile {
  id: string;
  filename: string;
  columns: string[];
  data: Record<string, string>[];
}

export const parseExcelFile = async (file: File): Promise<ParsedFile> => {
  const XLSX = await loadXlsx();

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Read all rows as arrays first
        const allRows = XLSX.utils.sheet_to_json<string[]>(worksheet, { header: 1 }) as string[][];
        
        if (allRows.length < 6) {
          reject(new Error('Excel file must have at least 6 rows (headers in rows 4-5, data starts from row 6)'));
          return;
        }
        
        // Row 4 (index 3) contains Blokk headers
        const blockHeaders = allRows[3] || [];
        // Row 5 (index 4) contains Navn and Klasse headers
        const mainHeaders = allRows[4] || [];
        
        // Combine headers: look for Navn, Klasse first, then Blokk columns
        const headers: string[] = [];
        const columnMapping: (number | null)[] = []; // Map to original column index
        
        // First pass: include all row-5 headers so users can remap in Oppsett.
        mainHeaders.forEach((header, idx) => {
          if (header) {
            headers.push(header);
            columnMapping.push(idx);
          }
        });
        
        // Second pass: get Blokk columns
        blockHeaders.forEach((header, idx) => {
          if (header && header.toLowerCase().includes('blokk')) {
            headers.push(header);
            columnMapping.push(idx);
          }
        });

        // Third pass: get Reserve columns
        blockHeaders.forEach((header, idx) => {
          if (header && header.toLowerCase().includes('reserve')) {
            headers.push(header);
            columnMapping.push(idx);
          }
        });

        // Fourth pass: get foreign language columns from row-4 labels.
        blockHeaders.forEach((header, idx) => {
          if (header && isForeignLanguageHeader(header)) {
            headers.push(header);
            columnMapping.push(idx);
          }
        });
        
        // Data starts from row 6 (index 5)
        const jsonData: Record<string, string>[] = [];
        for (let i = 5; i < allRows.length; i++) {
          const row = allRows[i];
          if (!row || row.every((cell) => !cell)) {
            // Skip empty rows
            continue;
          }
          
          const dataRow: Record<string, string> = {};
          headers.forEach((header, headerIdx) => {
            const colIdx = columnMapping[headerIdx];
            if (colIdx !== null && colIdx !== undefined) {
              dataRow[header] = (row[colIdx] || '').toString();
            }
          });
          
          jsonData.push(dataRow);
        }
        
        resolve({
          id: Math.random().toString(36).substr(2, 9),
          filename: file.name,
          columns: headers,
          data: jsonData,
        });
      } catch (error) {
        reject(error);
      }
    };
    
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
};

export interface ColumnMapping {
  [key: string]: string | null; // fileColumn -> standardField
}

const normalizeHeader = (header: string): string => {
  return header.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
};

const isBlokkMatVg2Header = (header: string): boolean => {
  return normalizeHeader(header) === 'blokkmatvg2';
};

const isReserveHeader = (header: string): boolean => {
  const normalized = normalizeHeader(header);
  return normalized.includes('reserve') || normalized === 'res';
};

const isForeignLanguageHeader = (header: string): boolean => {
  const normalized = header
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');

  return normalized.includes('fremmedsprak') || normalized.includes('fsp');
};

const isMath2PHeader = (header: string): boolean => {
  const normalized = header.trim().toLowerCase();
  return normalized === 'matematikk 2p' || normalized === '2p';
};

const isMathS1Header = (header: string): boolean => {
  const normalized = header.trim().toLowerCase();
  return normalized === 'matematikk s1' || normalized === 's1';
};

const isMathR1Header = (header: string): boolean => {
  const normalized = header.trim().toLowerCase();
  return normalized === 'matematikk r1' || normalized === 'r1';
};

export interface StandardField {
  studentId?: string;
  navn: string | null;
  klasse: string | null;
  blokkmatvg2: string | null;
  matematikk2p: string | null;
  matematikks1: string | null;
  matematikkr1: string | null;
  fremmedsprak: string | null;
  blokk1: string | null;
  blokk2: string | null;
  blokk3: string | null;
  blokk4: string | null;
  blokk5: string | null;
  blokk6: string | null;
  blokk7: string | null;
  blokk8: string | null;
  reserve: string | null;
}

export const getBlokkFields = (count: number): string[] => {
  const fields: string[] = [];
  for (let i = 1; i <= Math.min(count, 8); i++) {
    fields.push(`blokk${i}`);
  }
  return fields;
};

// Auto-detect column mappings based on column names
export const autoDetectMapping = (
  columns: string[],
  blokkCount: number = 4,
  rows: Record<string, string>[] = []
): ColumnMapping => {
  const mapping: ColumnMapping = {};
  
  columns.forEach((col) => {
    const colLower = col.toLowerCase();
    
    // Map Elevnavn to navn
    if (colLower.includes('elevnavn')) {
      mapping[col] = 'navn';
    }
    // Map Klasse to klasse
    else if (colLower === 'klasse') {
      mapping[col] = 'klasse';
    }
    // Map combined math choice column
    else if (isBlokkMatVg2Header(col)) {
      mapping[col] = 'blokkmatvg2';
    }
    // Map dedicated math columns
    else if (isMath2PHeader(col)) {
      mapping[col] = 'matematikk2p';
    }
    else if (isMathS1Header(col)) {
      mapping[col] = 'matematikks1';
    }
    else if (isMathR1Header(col)) {
      mapping[col] = 'matematikkr1';
    }
    // Map Reserve columns
    else if (isReserveHeader(col)) {
      mapping[col] = 'reserve';
    }
    // Map foreign language columns
    else if (isForeignLanguageHeader(col)) {
      mapping[col] = 'fremmedsprak';
    }
    // Map Blokk columns
    else if (colLower.includes('blokk')) {
      // Extract number from blokk header
      const match = colLower.match(/blokk(\d+)/);
      if (match) {
        const blokkNum = parseInt(match[1]);
        if (blokkNum > 0 && blokkNum <= blokkCount) {
          mapping[col] = `blokk${blokkNum}`;
        }
      }
    }
    else {
      mapping[col] = null;
    }
  });

  const alreadyMappedToBlokkMat = Object.values(mapping).includes('blokkmatvg2');
  if (!alreadyMappedToBlokkMat && rows.length > 0) {
    const mathTokenRegex = /(^|[^a-z0-9])(2p|s1|r1)([^a-z0-9]|$)/i;
    let bestColumn: string | null = null;
    let bestScore = 0;

    columns.forEach((col) => {
      const mappedField = mapping[col];
      const isReservedField = mappedField !== null && mappedField !== 'blokkmatvg2';
      if (isReservedField) {
        return;
      }

      const score = rows.reduce((acc, row) => {
        const value = (row[col] || '').toString().trim();
        if (!value) {
          return acc;
        }

        return acc + (mathTokenRegex.test(value) ? 1 : 0);
      }, 0);

      if (score > bestScore) {
        bestScore = score;
        bestColumn = col;
      }
    });

    if (bestColumn && bestScore > 0) {
      mapping[bestColumn] = 'blokkmatvg2';
    }
  }
  
  return mapping;
};

// Progress class to next year (1STA -> 2STA, 2STA -> 3STA, etc.)
const progressClass = (klasse: string): string => {
  // Match patterns like "1STA", "2VG2", "3STB", etc.
  const match = klasse.match(/^(\d+)(.+)$/);
  if (match) {
    const year = parseInt(match[1], 10);
    const suffix = match[2];
    return `${year + 1}${suffix}`;
  }
  return klasse; // Return unchanged if no match
};

export const mergeFiles = (
  files: ParsedFile[],
  mappings: Map<string, ColumnMapping>
): StandardField[] => {
  const merged: StandardField[] = [];
  
  files.forEach((file) => {
    const mapping = mappings.get(file.id) || {};
    
    file.data.forEach((row, rowIndex) => {
      const standardRow: StandardField = {
        studentId: `${file.id}:${rowIndex}`,
        navn: null,
        klasse: null,
        blokkmatvg2: null,
        matematikk2p: null,
        matematikks1: null,
        matematikkr1: null,
        fremmedsprak: null,
        blokk1: null,
        blokk2: null,
        blokk3: null,
        blokk4: null,
        blokk5: null,
        blokk6: null,
        blokk7: null,
        blokk8: null,
        reserve: null,
      };
      
      Object.entries(row).forEach(([fileColumn, value]) => {
        const standardField = mapping[fileColumn];
        if (standardField && value) {
          standardRow[standardField as keyof StandardField] = value;
        }
      });
      
      // Progress class to next year
      if (standardRow.klasse) {
        standardRow.klasse = progressClass(standardRow.klasse);
      }
      
      // Only add if at least navn is present
      if (standardRow.navn) {
        merged.push(standardRow);
      }
    });
  });
  
  return merged;
};

export interface SubjectCount {
  subject: string;
  count: number;
}

export interface StudentAssignmentChange {
  studentId: string;
  navn: string;
  klasse: string;
  subject: string;
  fromBlokk: number;
  toBlokk: number;
  reason: string;
  changedAt: string;
}

const parseSubjects = (value: string | null): string[] => {
  if (!value) {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((subject) => subject.trim())
    .filter((subject) => subject.length > 0);
};

const isSameSubject = (left: string, right: string): boolean => {
  return left.localeCompare(right, 'nb', { sensitivity: 'base' }) === 0;
};

const formatSubjects = (subjects: string[]): string | null => {
  if (subjects.length === 0) {
    return null;
  }

  return subjects.join(', ');
};

type BlokkField = `blokk${1 | 2 | 3 | 4 | 5 | 6 | 7 | 8}`;

const getBlokkField = (blokkNumber: number): BlokkField => {
  return `blokk${blokkNumber}` as BlokkField;
};

export const moveSubjectAssignmentsBetweenBlokker = (
  data: StandardField[],
  subject: string,
  fromBlokk: number,
  toBlokk: number,
  reason: string
): { updatedData: StandardField[]; changes: StudentAssignmentChange[] } => {
  if (fromBlokk === toBlokk) {
    return { updatedData: data, changes: [] };
  }

  if (fromBlokk < 1 || fromBlokk > 8 || toBlokk < 1 || toBlokk > 8) {
    return { updatedData: data, changes: [] };
  }

  const fromField = getBlokkField(fromBlokk);
  const toField = getBlokkField(toBlokk);
  const changes: StudentAssignmentChange[] = [];

  const updatedData = data.map((student, index) => {
    const fromValue = student[fromField];
    const toValue = student[toField];

    if (typeof fromValue !== 'string' || fromValue.trim().length === 0) {
      return student;
    }

    const fromSubjects = parseSubjects(fromValue);
    const hasSubjectInSource = fromSubjects.some((value) => isSameSubject(value, subject));
    if (!hasSubjectInSource) {
      return student;
    }

    const remainingSourceSubjects = fromSubjects.filter((value) => !isSameSubject(value, subject));
    const targetSubjects = parseSubjects(typeof toValue === 'string' ? toValue : null);
    const hasSubjectInTarget = targetSubjects.some((value) => isSameSubject(value, subject));
    const nextTargetSubjects = hasSubjectInTarget ? targetSubjects : [...targetSubjects, subject];

    const studentId = student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
    changes.push({
      studentId,
      navn: student.navn || 'Ukjent',
      klasse: student.klasse || 'Ingen klasse',
      subject,
      fromBlokk,
      toBlokk,
      reason,
      changedAt: new Date().toISOString(),
    });

    return {
      ...student,
      [fromField]: formatSubjects(remainingSourceSubjects),
      [toField]: formatSubjects(nextTargetSubjects),
    };
  });

  return { updatedData, changes };
};

export const swapSubjectAssignmentsBetweenBlokker = (
  data: StandardField[],
  subject: string,
  blokkA: number,
  blokkB: number,
  reason: string
): { updatedData: StandardField[]; changes: StudentAssignmentChange[] } => {
  if (blokkA === blokkB) {
    return { updatedData: data, changes: [] };
  }

  if (blokkA < 1 || blokkA > 8 || blokkB < 1 || blokkB > 8) {
    return { updatedData: data, changes: [] };
  }

  const fieldA = getBlokkField(blokkA);
  const fieldB = getBlokkField(blokkB);
  const changes: StudentAssignmentChange[] = [];

  const updatedData = data.map((student, index) => {
    const valueA = student[fieldA];
    const valueB = student[fieldB];

    const subjectsA = parseSubjects(typeof valueA === 'string' ? valueA : null);
    const subjectsB = parseSubjects(typeof valueB === 'string' ? valueB : null);
    const hasInA = subjectsA.some((value) => isSameSubject(value, subject));
    const hasInB = subjectsB.some((value) => isSameSubject(value, subject));

    if (hasInA === hasInB) {
      return student;
    }

    const nextA = subjectsA.filter((value) => !isSameSubject(value, subject));
    const nextB = subjectsB.filter((value) => !isSameSubject(value, subject));

    let fromBlokk = blokkA;
    let toBlokk = blokkB;
    if (hasInA && !hasInB) {
      nextB.push(subject);
    } else {
      nextA.push(subject);
      fromBlokk = blokkB;
      toBlokk = blokkA;
    }

    const studentId = student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
    changes.push({
      studentId,
      navn: student.navn || 'Ukjent',
      klasse: student.klasse || 'Ingen klasse',
      subject,
      fromBlokk,
      toBlokk,
      reason,
      changedAt: new Date().toISOString(),
    });

    return {
      ...student,
      [fieldA]: formatSubjects(nextA),
      [fieldB]: formatSubjects(nextB),
    };
  });

  return { updatedData, changes };
};

export const removeSubjectAssignmentsForStudents = (
  data: StandardField[],
  subject: string,
  studentIdsToUpdate: string[],
  reason: string
): { updatedData: StandardField[]; changes: StudentAssignmentChange[] } => {
  if (studentIdsToUpdate.length === 0) {
    return { updatedData: data, changes: [] };
  }

  const targetStudentIds = new Set(studentIdsToUpdate);
  const changes: StudentAssignmentChange[] = [];

  const updatedData = data.map((student, index) => {
    const studentId = student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
    if (!targetStudentIds.has(studentId)) {
      return student;
    }

    let didChange = false;
    const nextStudent: StandardField = { ...student };

    for (let blokkNumber = 1; blokkNumber <= 8; blokkNumber += 1) {
      const blokkField = getBlokkField(blokkNumber);
      const currentSubjects = parseSubjects(typeof nextStudent[blokkField] === 'string' ? nextStudent[blokkField] : null);
      const remainingSubjects = currentSubjects.filter((value) => !isSameSubject(value, subject));

      if (remainingSubjects.length === currentSubjects.length) {
        continue;
      }

      didChange = true;
      nextStudent[blokkField] = formatSubjects(remainingSubjects);

      changes.push({
        studentId,
        navn: student.navn || 'Ukjent',
        klasse: student.klasse || 'Ingen klasse',
        subject,
        fromBlokk: blokkNumber,
        toBlokk: 0,
        reason,
        changedAt: new Date().toISOString(),
      });
    }

    return didChange ? nextStudent : student;
  });

  return { updatedData, changes };
};

export const tallySubjects = (mergedData: StandardField[]): SubjectCount[] => {
  const subjectMap = new Map<string, number>();
  
  mergedData.forEach((row) => {
    [row.blokk1, row.blokk2, row.blokk3, row.blokk4].forEach((blokk) => {
      if (blokk) {
        // Split by comma or semicolon if multiple subjects
        const subjects = blokk.split(/[,;]/).map((s) => s.trim()).filter((s) => s);
        subjects.forEach((subject) => {
          subjectMap.set(subject, (subjectMap.get(subject) || 0) + 1);
        });
      }
    });
  });
  
  return Array.from(subjectMap.entries())
    .map(([subject, count]) => ({ subject, count }))
    .sort((a, b) => a.subject.localeCompare(b.subject));
};

const windows1252Map = new Map<number, number>([
  [0x20AC, 0x80],
  [0x201A, 0x82],
  [0x0192, 0x83],
  [0x201E, 0x84],
  [0x2026, 0x85],
  [0x2020, 0x86],
  [0x2021, 0x87],
  [0x02C6, 0x88],
  [0x2030, 0x89],
  [0x0160, 0x8A],
  [0x2039, 0x8B],
  [0x0152, 0x8C],
  [0x017D, 0x8E],
  [0x2018, 0x91],
  [0x2019, 0x92],
  [0x201C, 0x93],
  [0x201D, 0x94],
  [0x2022, 0x95],
  [0x2013, 0x96],
  [0x2014, 0x97],
  [0x02DC, 0x98],
  [0x2122, 0x99],
  [0x0161, 0x9A],
  [0x203A, 0x9B],
  [0x0153, 0x9C],
  [0x017E, 0x9E],
  [0x0178, 0x9F],
]);

const encodeWindows1252 = (value: string): Uint8Array => {
  const bytes: number[] = [];

  for (const character of value) {
    const codePoint = character.codePointAt(0);

    if (codePoint === undefined) {
      continue;
    }

    if (codePoint <= 0x7F || (codePoint >= 0xA0 && codePoint <= 0xFF)) {
      bytes.push(codePoint);
      continue;
    }

    const mappedByte = windows1252Map.get(codePoint);
    bytes.push(mappedByte ?? 0x3F);
  }

  return new Uint8Array(bytes);
};

const getBlockLetter = (blockNumber: number): string => {
  let result = '';
  let n = blockNumber;

  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }

  return result;
};

const mapBlockSubjectsToCodes = (subjectValue: string | null, blockNumber: number): string[] => {
  if (!subjectValue) {
    return [];
  }

  const blockLetter = getBlockLetter(blockNumber);
  return subjectValue
    .split(/[,;]/)
    .map((subject) => subject.trim())
    .filter((subject) => subject.length > 0)
    .map((subject) => `${mapSubjectToCode(subject)}${blockLetter}`);
};

const getCombinedBlockCodes = (row: StandardField): string => {
  return [row.blokk1, row.blokk2, row.blokk3, row.blokk4]
    .flatMap((subjectValue, index) => mapBlockSubjectsToCodes(subjectValue, index + 1))
    .join(',');
};

export const exportToExcel = async (mergedData: StandardField[], filename: string = 'merged_students.xlsx') => {
  const XLSX = await loadXlsx();

  const exportData = mergedData.map((row, index) => ({
    'Student ID': (1001 + index).toString(),
    'Navn': row.navn || '',
    'Klasse': row.klasse || '',
    'Blokk': getCombinedBlockCodes(row),
    'Reserve': row.reserve || '',
  }));
  
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Students');
  
  // Set column widths
  worksheet['!cols'] = [
    { wch: 12 },
    { wch: 20 },
    { wch: 12 },
    { wch: 32 },
    { wch: 24 },
  ];
  
  XLSX.writeFile(workbook, filename);
};

export const exportToExcelDetailed = async (
  mergedData: StandardField[],
  blokkCount: number,
  filename: string = 'merged_students_full.xlsx'
) => {
  const XLSX = await loadXlsx();
  const safeBlokkCount = Math.max(1, Math.min(8, blokkCount));

  const exportData = mergedData.map((row) => {
    const base: Record<string, string> = {
      Navn: row.navn || '',
      Klasse: row.klasse || '',
    };

    for (let i = 1; i <= safeBlokkCount; i++) {
      const key = `blokk${i}` as keyof StandardField;
      base[`Blokk ${i}`] = typeof row[key] === 'string' ? row[key] || '' : '';
    }

    base.Reserve = row.reserve || '';
    return base;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Students Full');

  const cols = [
    { wch: 24 },
    { wch: 12 },
    ...Array.from({ length: safeBlokkCount }, () => ({ wch: 24 })),
    { wch: 24 },
  ];
  worksheet['!cols'] = cols;

  XLSX.writeFile(workbook, filename);
};

/**
 * Export merged data as a tab-separated text file with student numbers and subject codes
 */
export const exportToTabText = (mergedData: StandardField[], filename: string = 'merged_students.txt') => {
  const mapReserveSubjects = (subjectValue: string | null): string => {
    if (!subjectValue) {
      return '';
    }

    return subjectValue
      .split(/[,;]/)
      .map((subject) => subject.trim())
      .filter((subject) => subject.length > 0)
      .map((subject) => `${subject} (${mapSubjectToCode(subject)})`)
      .join(',');
  };

  // Create rows with id, name, class and one combined subject column.
  const rows = mergedData.map((row, index) => {
    const studentNumber = (1001 + index).toString();
    const navn = row.navn || '';
    const klasse = row.klasse || '';
    const subjects = [
      ...mapBlockSubjectsToCodes(row.blokk1, 1),
      ...mapBlockSubjectsToCodes(row.blokk2, 2),
      ...mapBlockSubjectsToCodes(row.blokk3, 3),
      ...mapBlockSubjectsToCodes(row.blokk4, 4),
    ].join(',');
    const reserveSubjects = mapReserveSubjects(row.reserve);
    
    return [studentNumber, navn, klasse, subjects, reserveSubjects];
  });
  
  // Join each row with tabs
  const textContent = rows.map(row => row.join('\t')).join('\n');

  // Create a Blob and trigger download
  const encodedContent = encodeWindows1252(textContent);
  const encodedBuffer = new ArrayBuffer(encodedContent.byteLength);
  new Uint8Array(encodedBuffer).set(encodedContent);
  const blob = new Blob([encodedBuffer], { type: 'text/plain;charset=windows-1252' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};
