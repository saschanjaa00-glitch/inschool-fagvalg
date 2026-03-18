import { Fragment, useEffect, useMemo, useState } from 'react';
import type { SubjectCount, StandardField } from '../utils/excelUtils';
import { loadXlsx } from '../utils/excelUtils';
import {
  BLOKK_LABELS,
  DEFAULT_MAX_PER_SUBJECT,
  getActiveTotal,
  getBlokkNumber,
  getResolvedGroupsByTarget,
  getSettingsForSubject,
  makeGroup,
  sanitizeCount,
  shouldShowGroup,
  type BlokkLabel,
  type SubjectGroup,
  type SubjectSettingsByName,
  type StudentIdsByBlokk,
} from '../utils/subjectGroups';
import styles from './SubjectTally.module.css';

export type { SubjectSettingsByName } from '../utils/subjectGroups';

interface SubjectTallyProps {
  subjects: SubjectCount[];
  mergedData: StandardField[];
  subjectSettingsByName: SubjectSettingsByName;
  autoOpenSettingsToken?: number;
  onAutoOpenSettingsHandled?: (token: number) => void;
  onSaveSubjectSettingsByName: (values: SubjectSettingsByName) => void;
  onApplySubjectBlockMoves: (
    subject: string,
    operations: Array<
      | { type: 'move'; fromBlokk: number; toBlokk: number; reason: string }
      | { type: 'swap'; blokkA: number; blokkB: number; reason: string }
    >
  ) => void;
  onRemoveStudentsFromSubject: (subject: string, studentIds: string[], reason: string) => void;
  onOpenStudentCard: (studentId: string) => void;
}

interface ForeignLanguageOptionCount {
  label: string;
  count: number;
}

interface OptionStudent {
  studentId: string;
  navn: string;
  klasse: string;
}

interface OptionRow {
  label: string;
  count: number;
  students: OptionStudent[];
}

interface SubjectDraft {
  defaultMax: string;
  groupMaxBySlotKey: Record<string, string>;
}

const getDraftGroupMaxValue = (draft: SubjectDraft, slot: SubjectGroupSlot): string => {
  return draft.groupMaxBySlotKey[slot.slotKey] ?? draft.defaultMax;
};

const hasDraftGroupOverride = (draft: SubjectDraft, slot: SubjectGroupSlot): boolean => {
  return Object.prototype.hasOwnProperty.call(draft.groupMaxBySlotKey, slot.slotKey)
    && sanitizeCount(draft.groupMaxBySlotKey[slot.slotKey], sanitizeCount(draft.defaultMax, slot.max))
      !== sanitizeCount(draft.defaultMax, slot.max);
};

interface DeleteGroupConfirmState {
  subject: string;
  groupId: string;
  blokk: BlokkLabel;
  studentIds: string[];
}

interface SubjectGroupSlot {
  slotKey: string;
  groupId: string;
  label: string;
  blokk: BlokkLabel;
  max: number;
}

interface PdfGroupCell {
  count: number;
  enabled: boolean;
  overfilled: boolean;
}

const getPdfGroupColumnCount = (groupCount: number): number => {
  if (groupCount <= 1) {
    return 1;
  }

  if (groupCount === 2 || groupCount === 4) {
    return 2;
  }

  return 3;
};

const getPdfGroupRowCount = (groupCount: number): number => {
  if (groupCount <= 0) {
    return 1;
  }

  return Math.ceil(groupCount / getPdfGroupColumnCount(groupCount));
};

const parseSubjects = (value: string | null): string[] => {
  if (!value) {
    return [];
  }

  return value
    .split(/[,;]/)
    .map((subject) => subject.trim())
    .filter((subject) => subject.length > 0);
};

const parseForeignLanguageChoices = (value: string | null): string[] => {
  if (!value) {
    return [];
  }

  // Treat labels like "Tysk I+II, 2. år" as a single choice by removing year suffixes.
  const withoutYearSuffix = value.replace(/,\s*\d+\.?\s*(?:år|ar)\b/gi, '');

  return withoutYearSuffix
    .split(/[,;/]/)
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
};

const normalizeSubjectKey = (value: string): string => {
  return value.trim().toLocaleLowerCase('nb');
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

const MATH_OPTION_KEYS = ['R1', 'S1', '2P'] as const;
type MathOptionKey = (typeof MATH_OPTION_KEYS)[number];
const MATH_OPTION_DISPLAY: Record<MathOptionKey, string> = {
  R1: 'Matematikk R1',
  S1: 'Matematikk S1',
  '2P': 'Matematikk 2P',
};

const parseMathOptionsFromBlokkMat = (value: string | null): Set<MathOptionKey> => {
  const result = new Set<MathOptionKey>();
  if (!value) return result;
  value
    .split(/[,;/]/)
    .map((part) => part.trim().toUpperCase().replace(/\s+/g, ''))
    .filter((part) => part.length > 0)
    .forEach((part) => {
      if (part.includes('2P')) result.add('2P');
      if (part.includes('S1')) result.add('S1');
      if (part.includes('R1')) result.add('R1');
    });
  return result;
};

export const SubjectTally = ({
  subjects,
  mergedData,
  subjectSettingsByName,
  autoOpenSettingsToken,
  onAutoOpenSettingsHandled,
  onSaveSubjectSettingsByName,
  onApplySubjectBlockMoves,
  onRemoveStudentsFromSubject,
  onOpenStudentCard,
}: SubjectTallyProps) => {
  const [showOverfillModal, setShowOverfillModal] = useState(false);
  const [massUpdateMax, setMassUpdateMax] = useState(String(DEFAULT_MAX_PER_SUBJECT));
  const [draftsBySubject, setDraftsBySubject] = useState<Record<string, SubjectDraft>>({});
  const [expandedSettingsBySubject, setExpandedSettingsBySubject] = useState<Record<string, boolean>>({});
  const [copiedSubject, setCopiedSubject] = useState<string | null>(null);
  const [draggedSubject, setDraggedSubject] = useState<string | null>(null);
  const [draggedGroupId, setDraggedGroupId] = useState<string | null>(null);
  const [activeTrashSubject, setActiveTrashSubject] = useState<string | null>(null);
  const [deleteGroupConfirmState, setDeleteGroupConfirmState] = useState<DeleteGroupConfirmState | null>(null);
  const [isDeleteGroupConfirmArmed, setIsDeleteGroupConfirmArmed] = useState(false);
  const [expandedMathOption, setExpandedMathOption] = useState<string | null>(null);
  const [expandedForeignOption, setExpandedForeignOption] = useState<string | null>(null);
  const [showMath, setShowMath] = useState(false);

  const subjectStatsByKey = useMemo(() => {
    const stats = new Map<
      string,
      {
        breakdown: Record<BlokkLabel, number>;
        idsByBlokk: StudentIdsByBlokk;
      }
    >();

    const ensureStats = (subject: string) => {
      const key = normalizeSubjectKey(subject);
      if (!stats.has(key)) {
        stats.set(key, {
          breakdown: {
            'Blokk 1': 0,
            'Blokk 2': 0,
            'Blokk 3': 0,
            'Blokk 4': 0,
          },
          idsByBlokk: {
            'Blokk 1': [],
            'Blokk 2': [],
            'Blokk 3': [],
            'Blokk 4': [],
          },
        });
      }

      return stats.get(key)!;
    };

    subjects.forEach((item) => {
      ensureStats(item.subject);
    });
    Object.keys(subjectSettingsByName).forEach((subject) => {
      ensureStats(subject);
    });

    mergedData.forEach((student, index) => {
      const studentId = getStudentId(student, index);

      const subjectByBlokk: Array<{ label: BlokkLabel; value: string | null }> = [
        { label: 'Blokk 1', value: student.blokk1 },
        { label: 'Blokk 2', value: student.blokk2 },
        { label: 'Blokk 3', value: student.blokk3 },
        { label: 'Blokk 4', value: student.blokk4 },
      ];

      subjectByBlokk.forEach(({ label, value }) => {
        parseSubjects(value).forEach((subject) => {
          const subjectStats = ensureStats(subject);
          subjectStats.breakdown[label] += 1;
          subjectStats.idsByBlokk[label].push(studentId);
        });
      });

      if (showMath) {
        const mathOptions = parseMathOptionsFromBlokkMat(student.blokkmatvg2);
        mathOptions.forEach((key) => {
          const displayName = MATH_OPTION_DISPLAY[key];
          const mathStats = ensureStats(displayName);
          mathStats.breakdown['Blokk 1'] += 1;
          mathStats.idsByBlokk['Blokk 1'].push(studentId);
        });
      }
    });

    return stats;
  }, [mergedData, subjectSettingsByName, subjects, showMath]);

  const getBlokkBreakdown = (subject: string): Record<BlokkLabel, number> => {
    const entry = subjectStatsByKey.get(normalizeSubjectKey(subject));
    if (!entry) {
      return {
        'Blokk 1': 0,
        'Blokk 2': 0,
        'Blokk 3': 0,
        'Blokk 4': 0,
      };
    }

    return {
      'Blokk 1': entry.breakdown['Blokk 1'],
      'Blokk 2': entry.breakdown['Blokk 2'],
      'Blokk 3': entry.breakdown['Blokk 3'],
      'Blokk 4': entry.breakdown['Blokk 4'],
    };
  };

  const getStudentIdsByBlokk = (subject: string): StudentIdsByBlokk => {
    const entry = subjectStatsByKey.get(normalizeSubjectKey(subject));
    if (!entry) {
      return {
        'Blokk 1': [],
        'Blokk 2': [],
        'Blokk 3': [],
        'Blokk 4': [],
      };
    }

    return {
      'Blokk 1': [...entry.idsByBlokk['Blokk 1']],
      'Blokk 2': [...entry.idsByBlokk['Blokk 2']],
      'Blokk 3': [...entry.idsByBlokk['Blokk 3']],
      'Blokk 4': [...entry.idsByBlokk['Blokk 4']],
    };
  };

  const getResolvedForSubject = (subject: string, breakdown: Record<BlokkLabel, number>) => {
    const settings = getSettingsForSubject(subjectSettingsByName, subject, breakdown);
    const groups = settings.groups || [];
    const studentIdsByBlokk = getStudentIdsByBlokk(subject);
    const groupsByTarget = getResolvedGroupsByTarget(
      groups,
      studentIdsByBlokk,
      settings.groupStudentAssignments || {}
    );

    return {
      settings,
      groups,
      studentIdsByBlokk,
      groupsByTarget,
      activeTotal: getActiveTotal(groupsByTarget),
    };
  };

  const getGroupSlotsForSubject = (subject: string): SubjectGroupSlot[] => {
    const breakdown = getBlokkBreakdown(subject);
    const { groupsByTarget } = getResolvedForSubject(subject, breakdown);

    return BLOKK_LABELS.flatMap((blokk) =>
      groupsByTarget[blokk]
        .filter(shouldShowGroup)
        .map((group) => ({
          slotKey: `${blokk}|${group.label}`,
          groupId: group.id,
          label: group.label,
          blokk,
          max: group.max,
        }))
    );
  };

  const saveSubjectGroups = (subject: string, groups: SubjectGroup[], defaultMax?: number) => {
    const breakdown = getBlokkBreakdown(subject);
    const current = getSettingsForSubject(subjectSettingsByName, subject, breakdown);

    onSaveSubjectSettingsByName({
      ...subjectSettingsByName,
      [subject]: {
        defaultMax: defaultMax ?? current.defaultMax,
        groupStudentAssignments: { ...(current.groupStudentAssignments || {}) },
        groups,
      },
    });
  };

  const handleCopyTotal = async (subject: string, count: number) => {
    try {
      await navigator.clipboard.writeText(String(count));
      setCopiedSubject(subject);
      setTimeout(() => setCopiedSubject(null), 500);
    } catch (err) {
      console.error('Failed to copy:', err);
    }
  };

  const clearDraggedState = () => {
    setDraggedSubject(null);
    setDraggedGroupId(null);
    setActiveTrashSubject(null);
  };

  const closeDeleteGroupConfirm = () => {
    setDeleteGroupConfirmState(null);
    setIsDeleteGroupConfirmArmed(false);
  };

  const moveGroupToBlokk = (subject: string, groupId: string, targetBlokk: BlokkLabel) => {
    const breakdown = getBlokkBreakdown(subject);
    const { groups, groupsByTarget } = getResolvedForSubject(subject, breakdown);
    const allResolvedGroups = BLOKK_LABELS.flatMap((blokk) => groupsByTarget[blokk]);
    const movingGroup = allResolvedGroups.find((group) => group.id === groupId);
    const sourceBlokk = movingGroup?.blokk;
    const sourceEnabledGroups = sourceBlokk ? groupsByTarget[sourceBlokk].filter((group) => group.enabled) : [];
    const shouldMoveStudentsWithGroup = !!movingGroup
      && movingGroup.enabled
      && sourceEnabledGroups.length === 1
      && sourceEnabledGroups[0].id === movingGroup.id
      && sourceBlokk !== targetBlokk;

    const nextGroups = groups.map((group) => {
      if (group.id !== groupId) {
        return group;
      }

      return {
        ...group,
        blokk: targetBlokk,
      };
    });

    saveSubjectGroups(subject, nextGroups);

    if (shouldMoveStudentsWithGroup && sourceBlokk) {
      onApplySubjectBlockMoves(subject, [
        {
          type: 'move',
          fromBlokk: getBlokkNumber(sourceBlokk),
          toBlokk: getBlokkNumber(targetBlokk),
          reason: `Fagoversikt: flyttet gruppe ${movingGroup.label} fra ${sourceBlokk} til ${targetBlokk}`,
        },
      ]);
    }
  };

  const addExtraGroupToTarget = (subject: string, target: BlokkLabel) => {
    const breakdown = getBlokkBreakdown(subject);
    const { settings, groups } = getResolvedForSubject(subject, breakdown);

    const nextGroups = [
      ...groups,
      makeGroup(target, target, settings.defaultMax, true),
    ];

    saveSubjectGroups(subject, nextGroups, settings.defaultMax);
  };

  const removeDraggedGroup = (subject: string) => {
    if (draggedSubject !== subject || !draggedGroupId) {
      clearDraggedState();
      return;
    }

    const breakdown = getBlokkBreakdown(subject);
    const { groupsByTarget, groups } = getResolvedForSubject(subject, breakdown);

    const allResolved = BLOKK_LABELS.flatMap((blokk) => groupsByTarget[blokk]);
    const targetGroup = allResolved.find((group) => group.id === draggedGroupId);

    if (!targetGroup) {
      clearDraggedState();
      return;
    }

    const enabledGroupsInSameBlokk = groupsByTarget[targetGroup.blokk].filter((group) => group.enabled);
    const isLastEnabledGroupInBlokk = enabledGroupsInSameBlokk.length === 1
      && enabledGroupsInSameBlokk[0].id === targetGroup.id;

    if (targetGroup.allocatedCount > 0 && isLastEnabledGroupInBlokk) {
      setDeleteGroupConfirmState({
        subject,
        groupId: targetGroup.id,
        blokk: targetGroup.blokk,
        studentIds: targetGroup.allocatedStudentIds,
      });
      setIsDeleteGroupConfirmArmed(false);
      clearDraggedState();
      return;
    }

    if (targetGroup.allocatedCount > 0) {
      const nextGroups = groups.map((group) => {
        if (group.id !== draggedGroupId) {
          return group;
        }

        return {
          ...group,
          enabled: false,
        };
      });

      saveSubjectGroups(subject, nextGroups);
      clearDraggedState();
      return;
    }

    const nextGroups = groups.filter((group) => group.id !== draggedGroupId);
    saveSubjectGroups(subject, nextGroups);
    clearDraggedState();
  };

  const handleConfirmDeleteGroup = () => {
    if (!deleteGroupConfirmState) {
      return;
    }

    const { subject, groupId, blokk, studentIds } = deleteGroupConfirmState;
    const breakdown = getBlokkBreakdown(subject);
    const { settings, groups } = getResolvedForSubject(subject, breakdown);
    const nextGroups = groups.filter((group) => group.id !== groupId);
    const nextAssignments = Object.fromEntries(
      Object.entries(settings.groupStudentAssignments || {}).filter(([studentId, assignedGroupId]) => {
        return assignedGroupId !== groupId && !studentIds.includes(studentId);
      })
    );

    onSaveSubjectSettingsByName({
      ...subjectSettingsByName,
      [subject]: {
        ...settings,
        groups: nextGroups,
        groupStudentAssignments: nextAssignments,
      },
    });

    onRemoveStudentsFromSubject(
      subject,
      studentIds,
      `Fagoversikt: slettet siste gruppe i ${blokk}, fjernet ${subject} fra faget`
    );

    closeDeleteGroupConfirm();
  };

  const exportTable = async () => {
    const XLSX = await loadXlsx();

    const exportData = subjects.map((item) => {
      const breakdown = getBlokkBreakdown(item.subject);
      const { activeTotal } = getResolvedForSubject(item.subject, breakdown);

      return {
        Fag: item.subject,
        'Blokk 1': breakdown['Blokk 1'],
        'Blokk 2': breakdown['Blokk 2'],
        'Blokk 3': breakdown['Blokk 3'],
        'Blokk 4': breakdown['Blokk 4'],
        Totalt: activeTotal,
      };
    });

    const mathData = mathOptionRows.map((row) => ({
      Valg: row.label,
      Antall: row.count,
    }));

    const langData = foreignLanguageRows.map((row) => ({
      Valg: row.label,
      Antall: row.count,
    }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(exportData), 'Fagoversikt');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(mathData), 'Matematikkvalg');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(langData), 'Fremmedspråkvalg');
    XLSX.writeFile(workbook, 'subject_tally.xlsx');
  };

  const exportPdf = async () => {
    const [{ jsPDF }, { default: autoTable }] = await Promise.all([
      import('jspdf'),
      import('jspdf-autotable'),
    ]);

    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 40;
    const docWithTables = doc as typeof doc & { lastAutoTable?: { finalY: number } };

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Blokkoversikt', margin, margin);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.text(new Date().toLocaleString('nb-NO'), pageWidth - margin, margin, { align: 'right' });

    const pdfRows = subjectRows.map((row) => {
      const blockGroups = BLOKK_LABELS.reduce((acc, blokkLabel) => {
        acc[blokkLabel] = row.groupsByTarget[blokkLabel]
          .filter(shouldShowGroup)
          .map((entry) => ({
            count: entry.allocatedCount,
            enabled: entry.enabled,
            overfilled: entry.overfilled,
          }));
        return acc;
      }, {} as Record<BlokkLabel, PdfGroupCell[]>);

      const lineCount = Math.max(
        1,
        ...BLOKK_LABELS.map((blokkLabel) => getPdfGroupRowCount(blockGroups[blokkLabel].length))
      );
      const spacer = Array.from({ length: lineCount }, () => ' ').join('\n');

      return {
        subject: row.item.subject,
        blokk1: spacer,
        blokk2: spacer,
        blokk3: spacer,
        blokk4: spacer,
        total: String(row.activeTotal),
        blockGroups,
      };
    });

    autoTable(doc, {
      startY: margin + 18,
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 9,
        cellPadding: 6,
        lineColor: [220, 226, 234],
        lineWidth: 0.6,
        textColor: [34, 34, 34],
        valign: 'middle',
      },
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: [34, 34, 34],
        fontStyle: 'bold',
        halign: 'center',
      },
      columnStyles: {
        subject: { cellWidth: 190, halign: 'left' },
        blokk1: { cellWidth: 120, halign: 'center' },
        blokk2: { cellWidth: 120, halign: 'center' },
        blokk3: { cellWidth: 120, halign: 'center' },
        blokk4: { cellWidth: 120, halign: 'center' },
        total: { cellWidth: 65, halign: 'center', fontStyle: 'bold' },
      },
      columns: [
        { header: 'Fag', dataKey: 'subject' },
        { header: 'Blokk 1', dataKey: 'blokk1' },
        { header: 'Blokk 2', dataKey: 'blokk2' },
        { header: 'Blokk 3', dataKey: 'blokk3' },
        { header: 'Blokk 4', dataKey: 'blokk4' },
        { header: 'Totalt', dataKey: 'total' },
      ],
      body: pdfRows.map((row) => [
        row.subject,
        row.blokk1,
        row.blokk2,
        row.blokk3,
        row.blokk4,
        row.total,
      ]),
      didDrawCell: (data) => {
        if (data.section !== 'body') {
          return;
        }

        const dataKey = String(data.column.dataKey);
        const blokkLabel = (`Blokk ${dataKey.slice(-1)}`) as BlokkLabel;

        if (!['blokk1', 'blokk2', 'blokk3', 'blokk4'].includes(dataKey)) {
          return;
        }

        const groups = pdfRows[data.row.index]?.blockGroups[blokkLabel];
        if (!groups || groups.length === 0) {
          return;
        }

        const cols = getPdfGroupColumnCount(groups.length);
        const rows = getPdfGroupRowCount(groups.length);
        const gap = 4;
        const horizontalPadding = 8;
        const verticalPadding = 6;
        const availableWidth = data.cell.width - (horizontalPadding * 2) - (gap * (cols - 1));
        const availableHeight = data.cell.height - (verticalPadding * 2) - (gap * (rows - 1));
        const boxWidth = Math.max(20, Math.min(32, availableWidth / cols));
        const boxHeight = Math.max(16, Math.min(20, availableHeight / rows));
        const contentWidth = (boxWidth * cols) + (gap * (cols - 1));
        const contentHeight = (boxHeight * rows) + (gap * (rows - 1));
        const startX = data.cell.x + ((data.cell.width - contentWidth) / 2);
        const startY = data.cell.y + ((data.cell.height - contentHeight) / 2);

        groups.forEach((group, index) => {
          const columnIndex = index % cols;
          const rowIndex = Math.floor(index / cols);
          const x = startX + (columnIndex * (boxWidth + gap));
          const y = startY + (rowIndex * (boxHeight + gap));

          doc.setFillColor(group.enabled ? 245 : 250, group.enabled ? 248 : 250, group.enabled ? 252 : 250);
          doc.setDrawColor(group.overfilled ? 187 : 180, group.overfilled ? 74 : 187, group.overfilled ? 74 : 197);
          doc.roundedRect(x, y, boxWidth, boxHeight, 4, 4, 'FD');
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(9);
          doc.setTextColor(34, 34, 34);
          doc.text(String(group.count), x + (boxWidth / 2), y + (boxHeight / 2) + 3, { align: 'center' });
        });

        doc.setFont('helvetica', 'normal');
      },
    });

    let nextY = (docWithTables.lastAutoTable?.finalY || margin + 18) + 24;
    const ensurePageSpace = (requiredHeight: number) => {
      if (nextY + requiredHeight <= pageHeight - margin) {
        return;
      }

      doc.addPage();
      nextY = margin;
    };

    ensurePageSpace(120);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.text('Matematikkvalg', margin, nextY);
    autoTable(doc, {
      startY: nextY + 8,
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 9,
        cellPadding: 6,
        lineColor: [220, 226, 234],
        lineWidth: 0.6,
        textColor: [34, 34, 34],
      },
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: [34, 34, 34],
        fontStyle: 'bold',
      },
      columnStyles: {
        count: { cellWidth: 90, halign: 'center' },
      },
      columns: [
        { header: 'Fag', dataKey: 'label' },
        { header: 'Antall', dataKey: 'count' },
      ],
      body: mathOptionRows.map((row) => ({ label: row.label, count: String(row.count) })),
    });

    nextY = (docWithTables.lastAutoTable?.finalY || nextY) + 24;
    ensurePageSpace(140);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.text('Fremmedspråkvalg', margin, nextY);
    autoTable(doc, {
      startY: nextY + 8,
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 9,
        cellPadding: 6,
        lineColor: [220, 226, 234],
        lineWidth: 0.6,
        textColor: [34, 34, 34],
      },
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: [34, 34, 34],
        fontStyle: 'bold',
      },
      columnStyles: {
        count: { cellWidth: 90, halign: 'center' },
      },
      columns: [
        { header: 'Fag', dataKey: 'label' },
        { header: 'Antall', dataKey: 'count' },
      ],
      body: (foreignLanguageRows.length > 0 ? foreignLanguageRows : [{ label: 'Ingen registrerte valg', count: 0, students: [] }]).map((row) => ({
        label: row.label,
        count: String(row.count),
      })),
    });

    doc.save('blokkoversikt.pdf');
  };

  const exportStudentList = async (subject: string) => {
    const XLSX = await loadXlsx();

    const BLOKK_LABELS_ORDERED: BlokkLabel[] = ['Blokk 1', 'Blokk 2', 'Blokk 3', 'Blokk 4'];
    const studentIdsByBlokk = getStudentIdsByBlokk(subject);
    const studentById = new Map<string, StandardField>();
    mergedData.forEach((student, index) => {
      studentById.set(getStudentId(student, index), student);
    });

    const rows: { Fag: string; Blokk: string; Navn: string }[] = [];
    BLOKK_LABELS_ORDERED.forEach((blokkLabel) => {
      const ids = studentIdsByBlokk[blokkLabel];
      const names = ids
        .map((id) => studentById.get(id)?.navn || '')
        .filter(Boolean)
        .sort((a, b) => a.localeCompare(b, 'nb', { sensitivity: 'base' }));
      names.forEach((navn) => {
        rows.push({ Fag: subject, Blokk: blokkLabel, Navn: navn });
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Elevliste');
    XLSX.writeFile(workbook, `elevliste-${subject.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`);
  };

  const extractMathOptionsFromBlokkMat = (value: string | null): Set<'2P' | 'S1' | 'R1'> => {
    return parseMathOptionsFromBlokkMat(value);
  };

  const sortOptionStudents = (students: OptionStudent[]): OptionStudent[] => {
    return [...students].sort((left, right) => {
      const classCompare = left.klasse.localeCompare(right.klasse, 'nb', { sensitivity: 'base' });
      if (classCompare !== 0) {
        return classCompare;
      }

      return left.navn.localeCompare(right.navn, 'nb', { sensitivity: 'base' });
    });
  };

  const getStudentsForMathOption = (option: '2P' | 'S1' | 'R1'): OptionStudent[] => {
    const students = mergedData.reduce<OptionStudent[]>((acc, student, index) => {
      const selected = extractMathOptionsFromBlokkMat(student.blokkmatvg2);
      if (!selected.has(option)) {
        return acc;
      }

      acc.push({
        studentId: getStudentId(student, index),
        navn: student.navn || 'Ukjent',
        klasse: student.klasse || 'Ingen klasse',
      });
      return acc;
    }, []);

    return sortOptionStudents(students);
  };

  const mathOptionRows: OptionRow[] = useMemo(() => {
    const options: Array<{ label: string; key: '2P' | 'S1' | 'R1' }> = [
      { label: 'Matematikk 2P', key: '2P' },
      { label: 'Matematikk S1', key: 'S1' },
      { label: 'Matematikk R1', key: 'R1' },
    ];

    return options.map((option) => {
      const studentsForOption = getStudentsForMathOption(option.key);
      return {
        label: option.label,
        count: studentsForOption.length,
        students: studentsForOption,
      };
    });
  }, [mergedData]);

  const foreignLanguageRows: OptionRow[] = useMemo(() => {
    const byKey = new Map<string, { label: string; students: OptionStudent[] }>();

    mergedData.forEach((student, index) => {
      const rawValue = student.fremmedsprak;
      if (!rawValue) {
        return;
      }

      parseForeignLanguageChoices(rawValue).forEach((choice) => {
          const key = choice.toLowerCase();
          const existing = byKey.get(key);

          const optionStudent: OptionStudent = {
            studentId: getStudentId(student, index),
            navn: student.navn || 'Ukjent',
            klasse: student.klasse || 'Ingen klasse',
          };

          if (existing) {
            existing.students.push(optionStudent);
            return;
          }

          byKey.set(key, {
            label: choice,
            students: [optionStudent],
          });
        });
    });

    return Array.from(byKey.values())
      .map((entry) => ({
        label: entry.label,
        count: entry.students.length,
        students: sortOptionStudents(entry.students),
      }))
      .sort((left, right) => left.label.localeCompare(right.label, 'nb', { sensitivity: 'base' }));
  }, [mergedData]);

  const foreignLanguageOptionCounts: ForeignLanguageOptionCount[] = foreignLanguageRows.map((row) => ({
    label: row.label,
    count: row.count,
  }));

  const openOverfillModal = () => {
    const nextDrafts: Record<string, SubjectDraft> = {};
    const nextExpanded: Record<string, boolean> = {};

    subjects.forEach((item) => {
      const breakdown = getBlokkBreakdown(item.subject);
      const saved = getSettingsForSubject(subjectSettingsByName, item.subject, breakdown);
      const slots = getGroupSlotsForSubject(item.subject);
      const groupMaxBySlotKey = Object.fromEntries(
        slots
          .filter((slot) => sanitizeCount(slot.max, saved.defaultMax) !== sanitizeCount(saved.defaultMax))
          .map((slot) => [slot.slotKey, String(slot.max)])
      ) as Record<string, string>;

      nextDrafts[item.subject] = {
        defaultMax: String(saved.defaultMax),
        groupMaxBySlotKey,
      };
      nextExpanded[item.subject] = false;
    });

    setDraftsBySubject(nextDrafts);
    setExpandedSettingsBySubject(nextExpanded);
    setMassUpdateMax(String(DEFAULT_MAX_PER_SUBJECT));
    setShowOverfillModal(true);
  };

  useEffect(() => {
    if (!autoOpenSettingsToken || subjects.length === 0) {
      return;
    }

    openOverfillModal();
    onAutoOpenSettingsHandled?.(autoOpenSettingsToken);
    // We intentionally trigger on token changes from App to open once per load event.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoOpenSettingsToken]);

  const applyMassUpdate = () => {
    const safeValue = sanitizeCount(massUpdateMax);

    setDraftsBySubject((prev) => {
      const next = { ...prev };
      subjects.forEach((item) => {
        const draft = next[item.subject];
        if (!draft) {
          return;
        }
        next[item.subject] = {
          defaultMax: String(safeValue),
          groupMaxBySlotKey: Object.fromEntries(
            Object.entries(draft.groupMaxBySlotKey || {}).filter(([, value]) => {
              return sanitizeCount(value, safeValue) !== safeValue;
            })
          ),
        };
      });
      return next;
    });
  };

  const saveOverfillSettings = () => {
    const nextValues: SubjectSettingsByName = { ...subjectSettingsByName };

    subjects.forEach((item) => {
      const draft = draftsBySubject[item.subject];
      if (!draft) {
        return;
      }

      const breakdown = getBlokkBreakdown(item.subject);
      const current = getSettingsForSubject(subjectSettingsByName, item.subject, breakdown);
      const defaultMax = sanitizeCount(draft.defaultMax);
      const slotByGroupId = new Map<string, string>(
        getGroupSlotsForSubject(item.subject).map((slot) => [slot.groupId, slot.slotKey])
      );
      const groupMaxBySlotKey = draft.groupMaxBySlotKey || {};

      const nextGroups = (current.groups || []).map((group) => {
        const slotKey = slotByGroupId.get(group.id);
        if (slotKey && Object.prototype.hasOwnProperty.call(groupMaxBySlotKey, slotKey)) {
          return {
            ...group,
            max: sanitizeCount(groupMaxBySlotKey[slotKey], defaultMax),
          };
        }

        return {
          ...group,
          max: defaultMax,
        };
      });

      nextValues[item.subject] = {
        defaultMax,
        groupStudentAssignments: { ...(current.groupStudentAssignments || {}) },
        groups: nextGroups,
      };
    });

    onSaveSubjectSettingsByName(nextValues);
    setShowOverfillModal(false);
  };

  const subjectRows = useMemo(() => {
    return subjects.map((item) => {
      const breakdown = getBlokkBreakdown(item.subject);
      const resolved = getResolvedForSubject(item.subject, breakdown);

      return {
        item,
        breakdown,
        isMathOption: false,
        ...resolved,
      };
    });
  }, [subjects, mergedData, subjectSettingsByName]);

  const displaySubjectRows = useMemo(() => {
    if (!showMath) return subjectRows;

    const mathRows = MATH_OPTION_KEYS
      .map((key) => {
        const displayName = MATH_OPTION_DISPLAY[key];
        const breakdown = getBlokkBreakdown(displayName);
        const totalStudents = BLOKK_LABELS.reduce((s, b) => s + breakdown[b], 0);
        if (totalStudents === 0) return null;
        const resolved = getResolvedForSubject(displayName, breakdown);
        return { item: { subject: displayName, count: totalStudents }, breakdown, isMathOption: true, ...resolved };
      })
      .filter((r): r is NonNullable<typeof r> => r !== null);

    return [...subjectRows, ...mathRows]
      .sort((a, b) => a.item.subject.localeCompare(b.item.subject, 'nb', { sensitivity: 'base' }));
  }, [subjectRows, showMath, subjectStatsByKey]);

  if (subjects.length === 0) {
    return <div className={styles.empty}>Ingen fag funnet</div>;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.toolbar}>
        <label className={styles.mathToggleLabel}>
          <input
            type="checkbox"
            checked={showMath}
            onChange={(e) => setShowMath(e.target.checked)}
          />
          Vis matematikk
        </label>
        <button
          className={styles.settingsBtn}
          onClick={openOverfillModal}
          title="Overfyllingsinnstillinger"
        >
          Set maks elever / fag
        </button>
        <button
          className={styles.exportTableBtn}
          onClick={exportTable}
          title="Eksporter fagoversiktstabell"
        >
          Eksporter tabell
        </button>
        <button
          className={styles.exportTableBtn}
          onClick={() => { void exportPdf(); }}
          title="Eksporter blokkoversikt som PDF"
        >
          Eksporter PDF
        </button>
      </div>
      <table className={styles.table}>
        <thead>
          <tr>
            <th>Fag</th>
            <th>Blokk 1</th>
            <th>Blokk 2</th>
            <th>Blokk 3</th>
            <th>Blokk 4</th>
            {showMath && <th>Matte</th>}
            <th>Totalt</th>
            <th>Handlinger</th>
          </tr>
        </thead>
        <tbody>
          {displaySubjectRows.map((row) => {
            const renderGroupCell = (
              targetBlokk: BlokkLabel,
              keySuffix: string,
              titlePrefix?: string,
              options?: {
                entryFilter?: (entry: (typeof row.groupsByTarget)[BlokkLabel][number]) => boolean;
                dropTargetBlokk?: BlokkLabel;
                addTargetBlokk?: BlokkLabel;
                showAddButton?: boolean;
              }
            ) => {
              const entries = row.groupsByTarget[targetBlokk]
                .filter(shouldShowGroup)
                .filter((entry) => (options?.entryFilter ? options.entryFilter(entry) : true));
              const groupGridClassName = entries.length <= 1
                ? styles.groupCardsGridOne
                : (entries.length === 2 || entries.length === 4)
                  ? styles.groupCardsGridTwo
                  : styles.groupCardsGridThree;
              const blokkStudents = entries
                .filter((entry) => entry.enabled)
                .reduce((sum, entry) => sum + entry.allocatedCount, 0);
              const blokkSpaces = entries
                .filter((entry) => entry.enabled)
                .reduce((sum, entry) => sum + entry.max, 0);

              return (
                <td
                  key={`${row.item.subject}-${keySuffix}`}
                  title={`${titlePrefix || targetBlokk} (${blokkStudents} / ${blokkSpaces})`}
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={() => {
                    if (draggedSubject === row.item.subject && draggedGroupId) {
                      moveGroupToBlokk(row.item.subject, draggedGroupId, options?.dropTargetBlokk || targetBlokk);
                    }
                    clearDraggedState();
                  }}
                >
                  <div className={styles.groupStack}>
                    <div className={`${styles.groupCardsGrid} ${groupGridClassName}`.trim()}>
                      {entries.map((entry) => {
                        const isAtMax = entry.enabled && !entry.overfilled && entry.allocatedCount === entry.max;
                        return (
                          <div
                            key={`${row.item.subject}-${targetBlokk}-${entry.id}`}
                            className={`${styles.groupCard} ${entry.enabled ? styles.groupCardActive : styles.groupCardInactive} ${isAtMax ? styles.groupCardAtMax : ''} ${entry.overfilled ? styles.groupCardOverfilled : ''}`.trim()}
                            draggable={true}
                            onDragStart={(event) => {
                              event.dataTransfer.effectAllowed = 'move';
                              event.dataTransfer.setData('text/plain', `${row.item.subject}:${entry.id}`);
                              setDraggedSubject(row.item.subject);
                              setDraggedGroupId(entry.id);
                            }}
                            onDragEnd={clearDraggedState}
                            title={`${entry.label} (${entry.allocatedCount} / ${entry.max})`}
                          >
                            <span className={styles.groupCount}>{entry.allocatedCount}</span>
                          </div>
                        );
                      })}
                      {entries.length === 0 && <div className={styles.groupEmptySlot}>Tom</div>}
                    </div>
                    {(options?.showAddButton ?? true) && (
                      <button
                        type="button"
                        className={styles.groupAddButton}
                        onClick={(event) => {
                          event.stopPropagation();
                          addExtraGroupToTarget(row.item.subject, options?.addTargetBlokk || targetBlokk);
                        }}
                        title={`Legg til ny gruppe i ${titlePrefix || targetBlokk}`}
                        aria-label={`Legg til ny gruppe i ${titlePrefix || targetBlokk}`}
                      >
                        +
                      </button>
                    )}
                  </div>
                </td>
              );
            };

            return (
              <tr key={row.item.subject} className={styles.subjectRow}>
                <td className={styles.subjectNameCell}>{row.item.subject}</td>
                {BLOKK_LABELS.map((targetBlokk) => {
                  if (showMath && row.isMathOption && targetBlokk === 'Blokk 1') {
                    return renderGroupCell(targetBlokk, targetBlokk, undefined, {
                      entryFilter: (entry) => entry.sourceBlokk !== 'Blokk 1',
                    });
                  }
                  return renderGroupCell(targetBlokk, targetBlokk);
                })}
                {showMath && (
                  row.isMathOption
                    ? renderGroupCell('Blokk 1', 'Matte', 'Matte', {
                      entryFilter: (entry) => entry.sourceBlokk === 'Blokk 1',
                      dropTargetBlokk: 'Blokk 1',
                      addTargetBlokk: 'Blokk 1',
                    })
                    : <td className={styles.mathPlaceholderCell}>-</td>
                )}
                <td
                  className={styles.totalCell}
                  onDoubleClick={() => handleCopyTotal(row.item.subject, row.activeTotal)}
                  title="Dobbeltklikk for a kopiere"
                  style={{
                    cursor: 'pointer',
                    userSelect: 'none',
                    backgroundColor: copiedSubject === row.item.subject ? '#4CAF50' : undefined,
                    transition: 'background-color 0.5s ease-out',
                  }}
                >
                  {row.activeTotal}
                </td>
                <td>
                  <div
                    className={`${styles.trashDropZone} ${activeTrashSubject === row.item.subject ? styles.trashDropZoneActive : ''}`.trim()}
                    onDragOver={(event) => {
                      event.preventDefault();
                      if (activeTrashSubject !== row.item.subject) {
                        setActiveTrashSubject(row.item.subject);
                      }
                    }}
                    onDragEnter={(event) => {
                      event.preventDefault();
                      setActiveTrashSubject(row.item.subject);
                    }}
                    onDragLeave={() => {
                      if (activeTrashSubject === row.item.subject) {
                        setActiveTrashSubject(null);
                      }
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      removeDraggedGroup(row.item.subject);
                    }}
                    title="Dra en gruppe hit for a fjerne"
                    aria-label="Fjern gruppe"
                  >
                    <svg
                      className={styles.trashIcon}
                      viewBox="0 0 24 24"
                      aria-hidden="true"
                      focusable="false"
                    >
                      <path className={styles.trashLid} d="M9 3h6l1 2h4v2H4V5h4l1-2z" />
                      <path d="M7 7h10l-1 13H8L7 7z" />
                      <path d="M10 10v7" />
                      <path d="M14 10v7" />
                    </svg>
                  </div>
                  <button
                    type="button"
                    className={styles.exportListBtn}
                    onClick={() => { void exportStudentList(row.item.subject); }}
                    title={`Eksporter elevliste for ${row.item.subject}`}
                  >
                    Eksporter liste
                  </button>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>

      <h4 className={styles.subSectionTitle}>Matematikkvalg</h4>
      <table className={styles.mathTable}>
        <colgroup>
          <col />
          <col className={styles.mathCountColumn} />
        </colgroup>
        <thead>
          <tr>
            <th>Fag</th>
            <th className={styles.mathCountHeader}>Antall</th>
          </tr>
        </thead>
        <tbody>
          {mathOptionRows.flatMap((item) => {
            const rows = [
              <tr key={item.label}>
                <td>
                  <button
                    type="button"
                    className={styles.optionToggleBtn}
                    onClick={() => setExpandedMathOption((prev) => (prev === item.label ? null : item.label))}
                  >
                    {item.label}
                  </button>
                </td>
                <td className={styles.mathCountCell}>{item.count}</td>
              </tr>
            ];

            if (expandedMathOption === item.label) {
              rows.push(
                <tr key={`${item.label}-students`}>
                  <td colSpan={2} className={styles.optionStudentsCell}>
                    <div className={styles.optionStudentsList}>
                      {item.students.length === 0 ? (
                        <span className={styles.optionEmptyText}>Ingen elever</span>
                      ) : (
                        item.students.map((student) => (
                          <button
                            key={`${item.label}-${student.studentId}`}
                            type="button"
                            className={styles.optionStudentLink}
                            onClick={() => onOpenStudentCard(student.studentId)}
                          >
                            {student.klasse} - {student.navn}
                          </button>
                        ))
                      )}
                    </div>
                  </td>
                </tr>
              );
            }

            return rows;
          })}
        </tbody>
      </table>

      <h4 className={styles.subSectionTitle}>Fremmedspråkvalg</h4>
      <table className={styles.mathTable}>
        <colgroup>
          <col />
          <col className={styles.mathCountColumn} />
        </colgroup>
        <thead>
          <tr>
            <th>Fag</th>
            <th className={styles.mathCountHeader}>Antall</th>
          </tr>
        </thead>
        <tbody>
          {foreignLanguageOptionCounts.length === 0 ? (
            <tr>
              <td>Ingen registrerte valg</td>
              <td className={styles.mathCountCell}>0</td>
            </tr>
          ) : (
            foreignLanguageRows.flatMap((item) => {
              const rows = [
                <tr key={item.label}>
                  <td>
                    <button
                      type="button"
                      className={styles.optionToggleBtn}
                      onClick={() => setExpandedForeignOption((prev) => (prev === item.label ? null : item.label))}
                    >
                      {item.label}
                    </button>
                  </td>
                  <td className={styles.mathCountCell}>{item.count}</td>
                </tr>
              ];

              if (expandedForeignOption === item.label) {
                rows.push(
                  <tr key={`${item.label}-students`}>
                    <td colSpan={2} className={styles.optionStudentsCell}>
                      <div className={styles.optionStudentsList}>
                        {item.students.length === 0 ? (
                          <span className={styles.optionEmptyText}>Ingen elever</span>
                        ) : (
                          item.students.map((student) => (
                            <button
                              key={`${item.label}-${student.studentId}`}
                              type="button"
                              className={styles.optionStudentLink}
                              onClick={() => onOpenStudentCard(student.studentId)}
                            >
                              {student.klasse} - {student.navn}
                            </button>
                          ))
                        )}
                      </div>
                    </td>
                  </tr>
                );
              }

              return rows;
            })
          )}
        </tbody>
      </table>

      {showOverfillModal && (
        <div className={styles.modalOverlay} onClick={() => setShowOverfillModal(false)}>
          <div className={styles.modal} onClick={(event) => event.stopPropagation()}>
            <h4>Innstillinger - Sett maks antall elever per gruppe</h4>
            <div className={styles.massUpdateRow}>
              <label htmlFor="mass-update-max">Masseoppdater standard maks</label>
              <input
                id="mass-update-max"
                type="number"
                min="0"
                value={massUpdateMax}
                onChange={(event) => setMassUpdateMax(event.target.value)}
                className={styles.maxInput}
              />
              <button type="button" className={styles.modalSecondaryBtn} onClick={applyMassUpdate}>
                Bruk på alle
              </button>
            </div>

            <div className={styles.modalTableWrap}>
              <table className={styles.modalTable}>
                <colgroup>
                  <col className={styles.modalColSubject} />
                  <col className={styles.modalColMax} />
                </colgroup>
                <thead>
                  <tr>
                    <th>Fag</th>
                    <th>Standard maks</th>
                  </tr>
                </thead>
                <tbody>
                  {subjects.map((item) => {
                    const draft = draftsBySubject[item.subject];

                    if (!draft) {
                      return null;
                    }

                    const groupSlots = getGroupSlotsForSubject(item.subject);
                    const hasCustomGroupMax = groupSlots.some((slot) => {
                      return hasDraftGroupOverride(draft, slot);
                    });
                    const isExpanded = !!expandedSettingsBySubject[item.subject];

                    return (
                      <Fragment key={item.subject}>
                        <tr key={item.subject}>
                          <td>
                            <button
                              type="button"
                              className={styles.modalSubjectToggle}
                              onClick={() => {
                                setExpandedSettingsBySubject((prev) => ({
                                  ...prev,
                                  [item.subject]: !isExpanded,
                                }));
                              }}
                            >
                              <span className={styles.modalChevron}>{isExpanded ? '▼' : '▶'}</span>
                              <span>{item.subject}</span>
                            </button>
                          </td>
                          <td>
                            <div className={styles.standardMaxCell}>
                              <input
                                type="number"
                                min="0"
                                value={draft.defaultMax}
                                onChange={(event) => {
                                  const value = event.target.value;
                                  setDraftsBySubject((prev) => ({
                                    ...prev,
                                    [item.subject]: {
                                      defaultMax: value,
                                      groupMaxBySlotKey: Object.fromEntries(
                                        Object.entries(prev[item.subject]?.groupMaxBySlotKey || {}).filter(([, groupValue]) => {
                                          return sanitizeCount(groupValue, sanitizeCount(value, DEFAULT_MAX_PER_SUBJECT))
                                            !== sanitizeCount(value, DEFAULT_MAX_PER_SUBJECT);
                                        })
                                      ),
                                    },
                                  }));
                                }}
                                className={styles.maxInput}
                              />
                              {hasCustomGroupMax && (
                                <span className={styles.modalWarningWrap}>
                                  <span
                                    className={styles.modalWarningBadge}
                                    title="Én eller flere grupper avviker standard maks"
                                    aria-label="Én eller flere grupper avviker standard maks"
                                  >
                                    ⚠
                                  </span>
                                  <span className={styles.modalWarningText}>Én eller flere grupper avviker standard maks</span>
                                </span>
                              )}
                            </div>
                          </td>
                        </tr>
                        {isExpanded && groupSlots.map((slot) => {
                          const value = getDraftGroupMaxValue(draft, slot);
                          const isDifferentFromStandard = hasDraftGroupOverride(draft, slot);
                          return (
                            <tr
                              key={`${item.subject}-${slot.slotKey}`}
                              className={`${styles.modalGroupRow} ${isDifferentFromStandard ? styles.modalGroupRowWarning : ''}`.trim()}
                            >
                              <td className={styles.modalGroupLabelCell}>- {slot.label} ({slot.blokk})</td>
                              <td>
                                <input
                                  type="number"
                                  min="0"
                                  value={value}
                                  onChange={(event) => {
                                    const nextValue = event.target.value;
                                    setDraftsBySubject((prev) => ({
                                      ...prev,
                                      [item.subject]: (() => {
                                        const nextDefaultMax = prev[item.subject]?.defaultMax || draft.defaultMax;
                                        const nextOverrides = { ...(prev[item.subject]?.groupMaxBySlotKey || {}) };

                                        if (
                                          sanitizeCount(nextValue, sanitizeCount(nextDefaultMax, slot.max))
                                          === sanitizeCount(nextDefaultMax, slot.max)
                                        ) {
                                          delete nextOverrides[slot.slotKey];
                                        } else {
                                          nextOverrides[slot.slotKey] = nextValue;
                                        }

                                        return {
                                          defaultMax: nextDefaultMax,
                                          groupMaxBySlotKey: nextOverrides,
                                        };
                                      })(),
                                    }));
                                  }}
                                  className={`${styles.maxInput} ${isDifferentFromStandard ? styles.maxInputWarning : ''}`.trim()}
                                />
                              </td>
                            </tr>
                          );
                        })}
                      </Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className={styles.modalActions}>
              <button
                type="button"
                className={styles.modalSecondaryBtn}
                onClick={() => setShowOverfillModal(false)}
              >
                Avbryt
              </button>
              <button type="button" className={styles.modalPrimaryBtn} onClick={saveOverfillSettings}>
                Lagre
              </button>
            </div>
          </div>
        </div>
      )}

      {deleteGroupConfirmState && (
        <div
          className={styles.modalOverlay}
          onClick={() => {
            if (isDeleteGroupConfirmArmed) {
              setIsDeleteGroupConfirmArmed(false);
              return;
            }

            closeDeleteGroupConfirm();
          }}
        >
          <div
            className={styles.confirmModal}
            onClick={(event) => {
              event.stopPropagation();

              if (isDeleteGroupConfirmArmed) {
                setIsDeleteGroupConfirmArmed(false);
              }
            }}
          >
            <h4>Slett gruppe</h4>
            <p className={styles.confirmMessage}>
              Vil du slette denne gruppen? Elever som er tildelt gruppen vil fjernes fra faget.
            </p>
            <div className={styles.modalActions}>
              <button
                type="button"
                className={`${styles.modalSecondaryBtn} ${styles.confirmActionBtn}`}
                onClick={(event) => {
                  event.stopPropagation();
                  closeDeleteGroupConfirm();
                }}
              >
                Nei
              </button>
              <button
                type="button"
                className={`${styles.modalPrimaryBtn} ${styles.confirmActionBtn} ${
                  isDeleteGroupConfirmArmed ? styles.modalConfirmBtn : ''
                }`}
                onClick={(event) => {
                  event.stopPropagation();

                  if (isDeleteGroupConfirmArmed) {
                    handleConfirmDeleteGroup();
                    return;
                  }

                  setIsDeleteGroupConfirmArmed(true);
                }}
              >
                {isDeleteGroupConfirmArmed ? 'Bekreft' : 'Ja'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
