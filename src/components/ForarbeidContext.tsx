import { createContext, useContext, useState } from 'react';
import type { ReactNode } from 'react';

export type ForarbeidSubject = {
  id: string;
  name: string;
  groupCount: number;
  vgLevel: 'VG2' | 'VG3';
  blokkRestrictions: number[]; // Block numbers NOT allowed
};

export interface ForarbeidState {
  subjects: ForarbeidSubject[];
  blokkCount: number;
  addSubject: (subject: Partial<ForarbeidSubject>) => void;
  removeSubject: (id: string) => void;
  updateSubject: (id: string, changes: Partial<ForarbeidSubject>) => void;
  setBlokkCount: (count: number) => void;
}

const ForarbeidContext = createContext<ForarbeidState | undefined>(undefined);

export const useForarbeid = () => {
  const ctx = useContext(ForarbeidContext);
  if (!ctx) throw new Error('useForarbeid must be used within ForarbeidProvider');
  return ctx;
};


export interface ForarbeidProviderProps {
  children: ReactNode;
  initialSubjects?: ForarbeidSubject[];
  initialBlokkCount?: number;
}

export const ForarbeidProvider = ({ children, initialSubjects, initialBlokkCount }: ForarbeidProviderProps) => {
  const [subjects, setSubjects] = useState<ForarbeidSubject[]>(initialSubjects || []);
  const [blokkCount, setBlokkCount] = useState(initialBlokkCount ?? 4);

  const addSubject = (subject: Partial<ForarbeidSubject>) => {
    setSubjects((prev) => [
      ...prev,
      {
        id: Math.random().toString(36).slice(2),
        name: subject.name || '',
        groupCount: subject.groupCount || 1,
        vgLevel: subject.vgLevel || 'VG2',
        blokkRestrictions: subject.blokkRestrictions || [],
      },
    ]);
  };

  const removeSubject = (id: string) => {
    setSubjects((prev) => prev.filter((s) => s.id !== id));
  };

  const updateSubject = (id: string, changes: Partial<ForarbeidSubject>) => {
    setSubjects((prev) => prev.map((s) => (s.id === id ? { ...s, ...changes } : s)));
  };

  return (
    <ForarbeidContext.Provider value={{ subjects, blokkCount, addSubject, removeSubject, updateSubject, setBlokkCount }}>
      {children}
    </ForarbeidContext.Provider>
  );
};
