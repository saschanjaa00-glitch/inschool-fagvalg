import { useMemo, useState } from 'react';
import type { SubjectCount, StandardField } from '../utils/excelUtils';
import { loadXlsx } from '../utils/excelUtils';
import styles from './SubjectTally.module.css';

interface SubjectTallyProps {
  subjects: SubjectCount[];
  mergedData: StandardField[];
  subjectSettingsByName: SubjectSettingsByName;
  onSaveSubjectSettingsByName: (values: SubjectSettingsByName) => void;
  onApplySubjectBlockMoves: (
    subject: string,
    operations: Array<
      | { type: 'move'; fromBlokk: number; toBlokk: number; reason: string }
      | { type: 'swap'; blokkA: number; blokkB: number; reason: string }
    >
  ) => void;
}

interface MathOptionCount {
  label: string;
  count: number;
}

type BlokkLabel = 'Blokk 1' | 'Blokk 2' | 'Blokk 3' | 'Blokk 4';

interface SubjectGroup {
  id: string;
  blokk: BlokkLabel;
  sourceBlokk: BlokkLabel;
  enabled: boolean;
  max: number;
  createdAt: string;
}

interface SubjectSettings {
  defaultMax: number;
  groups?: SubjectGroup[];
  groupStudentAssignments?: Record<string, string>;
  // Legacy fields retained for backwards compatibility with persisted state.
  blokkMaxOverrides?: Partial<Record<BlokkLabel, number>>;
  blokkEnabled?: Partial<Record<BlokkLabel, boolean>>;
  blokkOrder?: BlokkLabel[];
  extraGroupCounts?: Partial<Record<BlokkLabel, number>>;
}

export type SubjectSettingsByName = Record<string, SubjectSettings>;

interface SubjectDraft {
  defaultMax: string;
}

interface ResolvedGroup extends SubjectGroup {
  label: string;
  allocatedCount: number;
  overfilled: boolean;
}

type StudentIdsByBlokk = Record<BlokkLabel, string[]>;

const DEFAULT_MAX_PER_SUBJECT = 30;
const BLOKK_LABELS: BlokkLabel[] = ['Blokk 1', 'Blokk 2', 'Blokk 3', 'Blokk 4'];

const buildDefaultSettings = (): SubjectSettings => ({
  defaultMax: DEFAULT_MAX_PER_SUBJECT,
  groups: [],
});

const normalizeOrder = (order?: BlokkLabel[]): BlokkLabel[] => {
  return BLOKK_LABELS.map((label, index) => {
    const candidate = order?.[index];
    return candidate && BLOKK_LABELS.includes(candidate) ? candidate : label;
  });
};

const sanitizeCount = (value: string | number | undefined, fallback: number = DEFAULT_MAX_PER_SUBJECT): number => {
  if (typeof value === 'number') {
    return Number.isNaN(value) ? fallback : Math.max(0, Math.floor(value));
  }

  const parsed = Number.parseInt(value || '', 10);
  return Number.isNaN(parsed) ? fallback : Math.max(0, parsed);
};

const makeGroupId = () => {
  return `group-${Math.random().toString(36).slice(2, 11)}`;
};

const makeGroup = (
  blokk: BlokkLabel,
  sourceBlokk: BlokkLabel,
  max: number,
  enabled: boolean,
  createdAt?: string
): SubjectGroup => ({
  id: makeGroupId(),
  blokk,
  sourceBlokk,
  enabled,
  max: sanitizeCount(max),
  createdAt: createdAt || new Date().toISOString(),
});

const sanitizeGroups = (groups: SubjectGroup[] | undefined, defaultMax: number): SubjectGroup[] => {
  if (!Array.isArray(groups)) {
    return [];
  }

  const seenIds = new Set<string>();

  return groups
    .filter((group) => BLOKK_LABELS.includes(group.blokk))
    .map((group) => {
      const id = group.id && !seenIds.has(group.id) ? group.id : makeGroupId();
      seenIds.add(id);

      return {
        id,
        blokk: group.blokk,
        sourceBlokk: BLOKK_LABELS.includes(group.sourceBlokk as BlokkLabel)
          ? (group.sourceBlokk as BlokkLabel)
          : group.blokk,
        enabled: group.enabled !== false,
        max: sanitizeCount(group.max, defaultMax),
        createdAt: group.createdAt || new Date().toISOString(),
      };
    });
};

const getBlokkNumber = (label: BlokkLabel): number => {
  return Number.parseInt(label.replace('Blokk ', ''), 10);
};

const buildLegacyGroups = (
  raw: SubjectSettings,
  defaultMax: number,
  blokkBreakdown: Record<BlokkLabel, number>
): SubjectGroup[] => {
  const groups: SubjectGroup[] = [];
  const placement = normalizeOrder(raw.blokkOrder);

  BLOKK_LABELS.forEach((sourceBlokk, sourceIndex) => {
    const targetBlokk = placement[sourceIndex] ?? sourceBlokk;
    const hasImportData = blokkBreakdown[sourceBlokk] > 0;
    const explicitEnabled = raw.blokkEnabled?.[sourceBlokk];
    const shouldCreate = hasImportData || typeof explicitEnabled === 'boolean';

    if (!shouldCreate) {
      return;
    }

    const max = sanitizeCount(raw.blokkMaxOverrides?.[sourceBlokk], defaultMax);
    const enabled = explicitEnabled ?? hasImportData;
    groups.push(makeGroup(targetBlokk, sourceBlokk, max, enabled));
  });

  BLOKK_LABELS.forEach((targetBlokk) => {
    const extras = Math.max(0, Math.floor(raw.extraGroupCounts?.[targetBlokk] ?? 0));
    for (let i = 0; i < extras; i += 1) {
      const max = sanitizeCount(raw.blokkMaxOverrides?.[targetBlokk], defaultMax);
      groups.push(makeGroup(targetBlokk, targetBlokk, max, true));
    }
  });

  return groups;
};

const buildImportGroups = (
  defaultMax: number,
  blokkBreakdown: Record<BlokkLabel, number>
): SubjectGroup[] => {
  const groups: SubjectGroup[] = [];

  BLOKK_LABELS.forEach((blokk) => {
    if (blokkBreakdown[blokk] > 0) {
      groups.push(makeGroup(blokk, blokk, defaultMax, true));
    }
  });

  return groups;
};

const getSettingsForSubject = (
  subjectSettingsByName: SubjectSettingsByName,
  subject: string,
  blokkBreakdown: Record<BlokkLabel, number>
): SubjectSettings => {
  const raw = subjectSettingsByName[subject];

  if (!raw) {
    const defaults = buildDefaultSettings();
    return {
      ...defaults,
      groupStudentAssignments: {},
      groups: buildImportGroups(defaults.defaultMax, blokkBreakdown),
    };
  }

  const defaultMax = sanitizeCount(raw.defaultMax);
  const explicitGroups = sanitizeGroups(raw.groups, defaultMax);
  const groupStudentAssignments = (raw.groupStudentAssignments && typeof raw.groupStudentAssignments === 'object')
    ? { ...raw.groupStudentAssignments }
    : {};

  if (explicitGroups.length > 0) {
    return {
      defaultMax,
      groupStudentAssignments,
      groups: explicitGroups,
    };
  }

  const hasLegacyConfig =
    !!raw.blokkOrder
    || !!raw.blokkEnabled
    || !!raw.blokkMaxOverrides
    || !!raw.extraGroupCounts;

  if (!hasLegacyConfig) {
    return {
      defaultMax,
      groupStudentAssignments,
      groups: buildImportGroups(defaultMax, blokkBreakdown),
    };
  }

  return {
    defaultMax,
    groupStudentAssignments,
    groups: buildLegacyGroups(raw, defaultMax, blokkBreakdown),
  };
};

const shouldShowGroup = (group: ResolvedGroup): boolean => {
  return group.allocatedCount > 0 || group.enabled;
};

const allocateCountsEvenlyAcrossGroups = (totalCount: number, groups: SubjectGroup[]): Record<string, number> => {
  const allocation: Record<string, number> = {};
  groups.forEach((group) => {
    allocation[group.id] = 0;
  });

  const enabledGroups = groups.filter((group) => group.enabled);
  if (enabledGroups.length === 0) {
    return allocation;
  }

  const safeTotal = Math.max(0, totalCount);
  const base = Math.floor(safeTotal / enabledGroups.length);
  const remainder = safeTotal % enabledGroups.length;

  enabledGroups.forEach((group, index) => {
    allocation[group.id] = base + (index < remainder ? 1 : 0);
  });

  return allocation;
};

const getResolvedGroupsByTarget = (
  groups: SubjectGroup[],
  blokkStudentIds: StudentIdsByBlokk,
  groupStudentAssignments: Record<string, string>
): Record<BlokkLabel, ResolvedGroup[]> => {
  const byTarget: Record<BlokkLabel, SubjectGroup[]> = {
    'Blokk 1': [],
    'Blokk 2': [],
    'Blokk 3': [],
    'Blokk 4': [],
  };

  groups.forEach((group) => {
    byTarget[group.blokk].push(group);
  });

  const resolvedByTarget: Record<BlokkLabel, ResolvedGroup[]> = {
    'Blokk 1': [],
    'Blokk 2': [],
    'Blokk 3': [],
    'Blokk 4': [],
  };

  BLOKK_LABELS.forEach((blokk) => {
    const sorted = [...byTarget[blokk]].sort((left, right) => {
      if (left.createdAt !== right.createdAt) {
        return left.createdAt.localeCompare(right.createdAt);
      }
      return left.id.localeCompare(right.id);
    });

    const enabledGroups = sorted.filter((group) => group.enabled);
    const allocation: Record<string, number> = {};
    sorted.forEach((group) => {
      allocation[group.id] = 0;
    });

    const studentIds = blokkStudentIds[blokk] || [];
    const unassignedStudentIds: string[] = [];

    studentIds.forEach((studentId) => {
      const assignedGroupId = groupStudentAssignments[studentId];
      if (!assignedGroupId) {
        unassignedStudentIds.push(studentId);
        return;
      }

      const assignedGroup = enabledGroups.find((group) => group.id === assignedGroupId);
      if (!assignedGroup) {
        unassignedStudentIds.push(studentId);
        return;
      }

      allocation[assignedGroup.id] = (allocation[assignedGroup.id] || 0) + 1;
    });

    const evenAllocation = allocateCountsEvenlyAcrossGroups(unassignedStudentIds.length, enabledGroups);
    Object.entries(evenAllocation).forEach(([groupId, count]) => {
      allocation[groupId] = (allocation[groupId] || 0) + count;
    });

    resolvedByTarget[blokk] = sorted.map((group, index) => {
      const count = allocation[group.id] ?? 0;
      return {
        ...group,
        label: `${getBlokkNumber(blokk)}-${index + 1}`,
        allocatedCount: count,
        overfilled: group.enabled && count > group.max,
      };
    });
  });

  return resolvedByTarget;
};

const getActiveTotal = (groupsByTarget: Record<BlokkLabel, ResolvedGroup[]>): number => {
  return BLOKK_LABELS.reduce((sum, blokk) => {
    const activeCount = groupsByTarget[blokk]
      .filter((group) => group.enabled)
      .reduce((groupSum, group) => groupSum + group.allocatedCount, 0);
    return sum + activeCount;
  }, 0);
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

const isSameSubject = (left: string, right: string): boolean => {
  return left.localeCompare(right, 'nb', { sensitivity: 'base' }) === 0;
};

const getStudentId = (student: StandardField, index: number): string => {
  return student.studentId || `${student.navn || 'ukjent'}:${student.klasse || 'ukjent'}:${index}`;
};

export const SubjectTally = ({
  subjects,
  mergedData,
  subjectSettingsByName,
  onSaveSubjectSettingsByName,
}: SubjectTallyProps) => {
  const [markOverfilled, setMarkOverfilled] = useState(false);
  const [showOverfillModal, setShowOverfillModal] = useState(false);
  const [massUpdateMax, setMassUpdateMax] = useState(String(DEFAULT_MAX_PER_SUBJECT));
  const [draftsBySubject, setDraftsBySubject] = useState<Record<string, SubjectDraft>>({});
  const [copiedSubject, setCopiedSubject] = useState<string | null>(null);
  const [draggedSubject, setDraggedSubject] = useState<string | null>(null);
  const [draggedGroupId, setDraggedGroupId] = useState<string | null>(null);

  const getBlokkBreakdown = (subject: string): Record<BlokkLabel, number> => {
    const blokkCounts: Record<BlokkLabel, number> = {
      'Blokk 1': 0,
      'Blokk 2': 0,
      'Blokk 3': 0,
      'Blokk 4': 0,
    };

    mergedData.forEach((student) => {
      if (student.blokk1?.split(/[,;]/).map((s) => s.trim()).includes(subject)) blokkCounts['Blokk 1'] += 1;
      if (student.blokk2?.split(/[,;]/).map((s) => s.trim()).includes(subject)) blokkCounts['Blokk 2'] += 1;
      if (student.blokk3?.split(/[,;]/).map((s) => s.trim()).includes(subject)) blokkCounts['Blokk 3'] += 1;
      if (student.blokk4?.split(/[,;]/).map((s) => s.trim()).includes(subject)) blokkCounts['Blokk 4'] += 1;
    });

    return blokkCounts;
  };

  const getStudentIdsByBlokk = (subject: string): StudentIdsByBlokk => {
    const idsByBlokk: StudentIdsByBlokk = {
      'Blokk 1': [],
      'Blokk 2': [],
      'Blokk 3': [],
      'Blokk 4': [],
    };

    mergedData.forEach((student, index) => {
      const studentId = getStudentId(student, index);

      const subjectByBlokk: Array<{ label: BlokkLabel; value: string | null }> = [
        { label: 'Blokk 1', value: student.blokk1 },
        { label: 'Blokk 2', value: student.blokk2 },
        { label: 'Blokk 3', value: student.blokk3 },
        { label: 'Blokk 4', value: student.blokk4 },
      ];

      subjectByBlokk.forEach(({ label, value }) => {
        const subjects = parseSubjects(value);
        if (subjects.some((item) => isSameSubject(item, subject))) {
          idsByBlokk[label].push(studentId);
        }
      });
    });

    return idsByBlokk;
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
  };

  const moveGroupToBlokk = (subject: string, groupId: string, targetBlokk: BlokkLabel) => {
    const breakdown = getBlokkBreakdown(subject);
    const { groups } = getResolvedForSubject(subject, breakdown);

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

  const exportTable = async () => {
    const XLSX = await loadXlsx();

    const exportData = subjects.map((item) => {
      const breakdown = getBlokkBreakdown(item.subject);
      const { groupsByTarget, activeTotal } = getResolvedForSubject(item.subject, breakdown);

      const formatBlokkCell = (entries: ResolvedGroup[]) => {
        const visibleEntries = entries.filter(shouldShowGroup);

        if (visibleEntries.length === 0) {
          return '';
        }

        const activeCount = visibleEntries
          .filter((entry) => entry.enabled)
          .reduce((sum, entry) => sum + entry.allocatedCount, 0);

        const labels = visibleEntries.map((entry) => entry.label).join(', ');
        return `${activeCount} (${labels})`;
      };

      return {
        Fag: item.subject,
        Blokk1: formatBlokkCell(groupsByTarget['Blokk 1']),
        Blokk2: formatBlokkCell(groupsByTarget['Blokk 2']),
        Blokk3: formatBlokkCell(groupsByTarget['Blokk 3']),
        Blokk4: formatBlokkCell(groupsByTarget['Blokk 4']),
        TotaltAktive: activeTotal,
        TotaltOriginalt: item.count,
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Fagoversikt');
    XLSX.writeFile(workbook, 'subject_tally.xlsx');
  };

  const extractMathOptionsFromBlokkMat = (value: string | null): Set<'2P' | 'S1' | 'R1'> => {
    const selected = new Set<'2P' | 'S1' | 'R1'>();

    if (!value) {
      return selected;
    }

    value
      .split(/[,;/]/)
      .map((part) => part.trim().toUpperCase().replace(/\s+/g, ''))
      .filter((part) => part.length > 0)
      .forEach((part) => {
        if (part.includes('2P')) {
          selected.add('2P');
        }
        if (part.includes('S1')) {
          selected.add('S1');
        }
        if (part.includes('R1')) {
          selected.add('R1');
        }
      });

    return selected;
  };

  const countMathOption = (option: '2P' | 'S1' | 'R1'): number => {
    return mergedData.reduce((count, student) => {
      const selected = extractMathOptionsFromBlokkMat(student.blokkmatvg2);
      return count + (selected.has(option) ? 1 : 0);
    }, 0);
  };

  const mathOptionCounts: MathOptionCount[] = [
    {
      label: 'Matematikk 2P',
      count: countMathOption('2P'),
    },
    {
      label: 'Matematikk S1',
      count: countMathOption('S1'),
    },
    {
      label: 'Matematikk R1',
      count: countMathOption('R1'),
    },
  ];

  const openOverfillModal = () => {
    const nextDrafts: Record<string, SubjectDraft> = {};

    subjects.forEach((item) => {
      const breakdown = getBlokkBreakdown(item.subject);
      const saved = getSettingsForSubject(subjectSettingsByName, item.subject, breakdown);
      nextDrafts[item.subject] = {
        defaultMax: String(saved.defaultMax),
      };
    });

    setDraftsBySubject(nextDrafts);
    setMassUpdateMax(String(DEFAULT_MAX_PER_SUBJECT));
    setShowOverfillModal(true);
  };

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

      // Keep individual group max values intact. Only replace values equal to prior default.
      const nextGroups = (current.groups || []).map((group) => {
        if (group.max === current.defaultMax) {
          return {
            ...group,
            max: defaultMax,
          };
        }
        return group;
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
      const overfilled = BLOKK_LABELS.some((blokk) => {
        return resolved.groupsByTarget[blokk].some((group) => group.overfilled);
      });

      return {
        item,
        breakdown,
        ...resolved,
        overfilled,
      };
    });
  }, [subjects, mergedData, subjectSettingsByName]);

  if (subjects.length === 0) {
    return <div className={styles.empty}>Ingen fag funnet</div>;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.toolbar}>
        <button
          className={styles.exportTableBtn}
          onClick={exportTable}
          title="Eksporter fagoversiktstabell"
        >
          Eksport tabell
        </button>
        <button
          className={`${styles.overfillBtn} ${markOverfilled ? styles.overfillBtnActive : ''}`.trim()}
          onClick={() => setMarkOverfilled((prev) => !prev)}
          title="Veksle fremheving av overfylte"
        >
          Merk overfylte
        </button>
        <button
          className={styles.settingsBtn}
          onClick={openOverfillModal}
          title="Overfyllingsinnstillinger"
        >
          Innstillinger
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
            <th>Totalt</th>
            <th>Handlinger</th>
          </tr>
        </thead>
        <tbody>
          {subjectRows.map((row) => {
            return (
              <tr key={row.item.subject} className={styles.subjectRow}>
                <td className={`${styles.subjectNameCell} ${markOverfilled && row.overfilled ? styles.overfilledSubject : ''}`.trim()}>{row.item.subject}</td>
                {BLOKK_LABELS.map((targetBlokk) => {
                  const entries = row.groupsByTarget[targetBlokk].filter(shouldShowGroup);
                  const blokkStudents = entries
                    .filter((entry) => entry.enabled)
                    .reduce((sum, entry) => sum + entry.allocatedCount, 0);
                  const blokkSpaces = entries
                    .filter((entry) => entry.enabled)
                    .reduce((sum, entry) => sum + entry.max, 0);

                  return (
                    <td
                      key={`${row.item.subject}-${targetBlokk}`}
                      title={`${targetBlokk} (${blokkStudents} / ${blokkSpaces})`}
                      onDragOver={(event) => event.preventDefault()}
                      onDrop={() => {
                        if (draggedSubject === row.item.subject && draggedGroupId) {
                          moveGroupToBlokk(row.item.subject, draggedGroupId, targetBlokk);
                        }
                        clearDraggedState();
                      }}
                    >
                      <div className={styles.groupStack}>
                        {entries.map((entry) => {
                          return (
                            <div
                              key={`${row.item.subject}-${targetBlokk}-${entry.id}`}
                              className={`${styles.groupCard} ${entry.enabled ? styles.groupCardActive : styles.groupCardInactive} ${entry.overfilled ? styles.groupCardOverfilled : ''}`.trim()}
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
                        <button
                          type="button"
                          className={styles.groupAddButton}
                          onClick={(event) => {
                            event.stopPropagation();
                            addExtraGroupToTarget(row.item.subject, targetBlokk);
                          }}
                          title={`Legg til ny gruppe i ${targetBlokk}`}
                          aria-label={`Legg til ny gruppe i ${targetBlokk}`}
                        >
                          +
                        </button>
                      </div>
                    </td>
                  );
                })}
                <td
                  className={`${styles.totalCell} ${markOverfilled && row.overfilled ? styles.totalCellOverfilled : ''}`.trim()}
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
                    className={styles.trashDropZone}
                    onDragOver={(event) => event.preventDefault()}
                    onDrop={(event) => {
                      event.preventDefault();
                      removeDraggedGroup(row.item.subject);
                    }}
                    title="Dra en gruppe hit for a fjerne"
                  >
                    X
                  </div>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>

      <h4 className={styles.subSectionTitle}>Matematikkvalg</h4>
      <table className={styles.mathTable}>
        <thead>
          <tr>
            <th>Fag</th>
            <th>Antall</th>
          </tr>
        </thead>
        <tbody>
          {mathOptionCounts.map((item) => (
            <tr key={item.label}>
              <td>{item.label}</td>
              <td className={styles.mathCountCell}>{item.count}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {showOverfillModal && (
        <div className={styles.modalOverlay} onClick={() => setShowOverfillModal(false)}>
          <div className={styles.modal} onClick={(event) => event.stopPropagation()}>
            <h4>Merk overfylt</h4>
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
                Bruk pa alle
              </button>
            </div>

            <div className={styles.modalTableWrap}>
              <table className={styles.modalTable}>
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

                    return (
                      <tr key={item.subject}>
                        <td>{item.subject}</td>
                        <td>
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
                                },
                              }));
                            }}
                            className={styles.maxInput}
                          />
                        </td>
                      </tr>
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
    </div>
  );
};
