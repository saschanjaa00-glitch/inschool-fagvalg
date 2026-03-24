import React, { useState } from 'react';
import { useForarbeid } from './ForarbeidContext';
import styles from './ForarbeidView.module.css';
import { findOptimalBlockAssignment } from '../utils/forarbeidOptimizer';
import { runForarbeidParallel } from '../utils/forarbeidOptimizer.workerPool';

const ForarbeidView: React.FC = () => {
  const { subjects, blokkCount, addSubject, removeSubject, updateSubject, setBlokkCount } = useForarbeid();
  const [result, setResult] = useState<ReturnType<typeof findOptimalBlockAssignment> | null>(null);
  const [calculating, setCalculating] = useState(false);
  const [progress, setProgress] = useState<{ current: number; total: number } | null>(null);

  const handleCalculate = async () => {
    setCalculating(true);
    setProgress(null);
    try {
      // Try parallel worker pool (no progress yet)
      const res = await runForarbeidParallel(subjects, blokkCount);
      setResult(res);
    } catch (e) {
      // Fallback to single-threaded with progress
      setResult(findOptimalBlockAssignment(subjects, blokkCount, setProgress));
    } finally {
      setCalculating(false);
    }
  };

  return (
    <div className="content-container" style={{ marginTop: 24 }}>
      <div className="headerRow">
        <h2 className="title">Forarbeid: Kombinasjonsutforsker</h2>
        <button className="exportTableBtn" onClick={() => addSubject({ name: '', groupCount: 1, vgLevel: 'VG2', blokkRestrictions: [] })}>
          + Legg til fag
        </button>
      </div>
      <div className="blokkCountSelector">
        <label>
          Antall blokker:
          <input
            type="number"
            min={1}
            max={8}
            value={blokkCount}
            onChange={e => setBlokkCount(Number(e.target.value))}
            className="blokkCountInput"
            style={{ marginLeft: 8 }}
          />
        </label>
      </div>
      <table className="subjectTable" style={{ width: '100%', marginTop: 16, marginBottom: 16 }}>
        <thead>
          <tr>
            <th>Fag</th>
            <th>Grupper</th>
            <th>VG-nivå</th>
            <th>Blokk-restriksjoner</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {subjects.map((subject, idx) => (
            <tr key={subject.id}>
              <td>
                <input
                  value={subject.name}
                  onChange={e => updateSubject(subject.id, { name: e.target.value })}
                  placeholder="Fagnavn"
                  style={{ width: 120 }}
                />
              </td>
              <td>
                <input
                  type="number"
                  min={1}
                  max={8}
                  value={subject.groupCount}
                  onChange={e => updateSubject(subject.id, { groupCount: Number(e.target.value) })}
                  style={{ width: 50 }}
                />
              </td>
              <td>
                <select
                  value={subject.vgLevel}
                  onChange={e => updateSubject(subject.id, { vgLevel: e.target.value as 'VG2' | 'VG3' })}
                >
                  <option value="VG2">VG2</option>
                  <option value="VG3">VG3</option>
                </select>
              </td>
              <td>
                {Array.from({ length: blokkCount }, (_, i) => i + 1).map(blokk => (
                  <label key={blokk} style={{ marginRight: 8, fontSize: 12 }}>
                    <input
                      type="checkbox"
                      checked={subject.blokkRestrictions.includes(blokk)}
                      onChange={e => {
                        const next = e.target.checked
                          ? [...subject.blokkRestrictions, blokk]
                          : subject.blokkRestrictions.filter(b => b !== blokk);
                        updateSubject(subject.id, { blokkRestrictions: next });
                      }}
                    />
                    Blokk {blokk}
                  </label>
                ))}
              </td>
              <td>
                <button className="exportListBtn" onClick={() => removeSubject(subject.id)} title="Fjern fag">🗑️</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{ marginTop: 24 }}>
        <button className="exportTableBtn" onClick={handleCalculate} disabled={subjects.length < 3 || calculating}>
          {calculating ? 'Beregner...' : 'Kjør optimalisering'}
        </button>
        {calculating && progress && (
          <div style={{ marginTop: 12, fontSize: 14 }}>
            Fremdrift: {progress.current.toLocaleString()} / {progress.total.toLocaleString()} ({((progress.current/progress.total)*100).toFixed(1)}%)
          </div>
        )}
      </div>
      {result && (
        <div className="empty" style={{ marginTop: 32 }}>
          <h3 style={{ marginTop: 0 }}>Resultat</h3>
          <div>Antall mulige 3-fagskombinasjoner: <strong>{result.maxCombinations}</strong></div>
          <table style={{ width: '100%', marginTop: 16, marginBottom: 16 }}>
            <thead>
              <tr>
                <th>Fag</th>
                <th>Blokk</th>
              </tr>
            </thead>
            <tbody>
              {subjects.map(s => (
                <tr key={s.id}>
                  <td>{s.name || <em>Uten navn</em>}</td>
                  <td>{result.assignment[s.id] ?? <em>Ikke plassert</em>}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <h4>Kombinasjonsdetaljer</h4>
          <ul>
            {result.details.map((d, i) => (
              <li key={i}>
                Blokk {d.blocks.join(' + ')}: {d.count} kombinasjoner
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
};

export default ForarbeidView;
