import React, { useState } from 'react';
import { SchoolData, TestPhase, ReportConfig, PhaseStats } from '../types';

interface ManualEntryFormProps {
  config: ReportConfig;
  onSubmit: (data: SchoolData, phase: TestPhase) => void;
  onCancel: () => void;
}

export default function ManualEntryForm({ config, onSubmit, onCancel }: ManualEntryFormProps) {
  const [district, setDistrict] = useState('');
  const [mandal, setMandal] = useState('');
  const [schoolName, setSchoolName] = useState('');
  const [targetPhase, setTargetPhase] = useState<TestPhase>('Baseline');

  const emptyPhase = (): PhaseStats => ({ KNOW: 0, READ: 0, SPELL: 0, CWR: 0, CWS: 0, Total: 0 });

  const [gradesData, setGradesData] = useState<any>({
    3: { assessedCount: 0, baseline: emptyPhase(), midline: emptyPhase(), endline: emptyPhase() },
    4: { assessedCount: 0, baseline: emptyPhase(), midline: emptyPhase(), endline: emptyPhase() },
    5: { assessedCount: 0, baseline: emptyPhase(), midline: emptyPhase(), endline: emptyPhase() }
  });

  const handlePaste = (e: React.ClipboardEvent<HTMLInputElement>, grade: number, phase: 'baseline' | 'midline' | 'endline') => {
    const pastedData = e.clipboardData.getData('Text');
    const values = pastedData.split(/[\t\n,]/).map(v => v.trim()).filter(v => v !== '');
    
    if (values.length >= 5) {
      e.preventDefault();
      const KNOW = Number(values[0]) || 0;
      const READ = Number(values[1]) || 0;
      const SPELL = Number(values[2]) || 0;
      const CWR = Number(values[3]) || 0;
      const CWS = Number(values[4]) || 0;
      const Total = KNOW + READ + SPELL + CWR + CWS;
      
      setGradesData((prev: any) => ({
        ...prev,
        [grade]: {
          ...prev[grade],
          [phase]: { KNOW, READ, SPELL, CWR, CWS, Total }
        }
      }));
    }
  };

  const handleParamChange = (grade: number, phase: 'baseline' | 'midline' | 'endline', param: keyof PhaseStats, value: string) => {
    setGradesData((prev: any) => {
      const numValue = Number(value) || 0;
      const newPhaseData = { ...prev[grade][phase], [param]: numValue };
      newPhaseData.Total = newPhaseData.KNOW + newPhaseData.READ + newPhaseData.SPELL + newPhaseData.CWR + newPhaseData.CWS;
      return {
        ...prev,
        [grade]: {
          ...prev[grade],
          [phase]: newPhaseData
        }
      };
    });
  };

  const handleCountChange = (grade: number, value: string) => {
    setGradesData((prev: any) => ({
      ...prev,
      [grade]: {
        ...prev[grade],
        assessedCount: Number(value) || 0
      }
    }));
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!schoolName) {
      alert("Please enter a School Name");
      return;
    }
    const schoolData: SchoolData = {
      District: district,
      Mandal: mandal,
      School_Name: schoolName,
      classes: {
        3: gradesData[3],
        4: gradesData[4],
        5: gradesData[5]
      }
    };
    onSubmit(schoolData, targetPhase);

    setSchoolName('');
    setGradesData({
      3: { assessedCount: 0, baseline: emptyPhase(), midline: emptyPhase(), endline: emptyPhase() },
      4: { assessedCount: 0, baseline: emptyPhase(), midline: emptyPhase(), endline: emptyPhase() },
      5: { assessedCount: 0, baseline: emptyPhase(), midline: emptyPhase(), endline: emptyPhase() }
    });

    alert(`Report for ${schoolName} generated! You can now enter the next school.`);
  };

  const paramKeys: Array<keyof ReportConfig['parameters']> = ['KNOW', 'READ', 'SPELL', 'CWR', 'CWS'];

  const renderPhaseInputs = (grade: number, phase: 'baseline' | 'midline' | 'endline') => {
    return (
      <div className="mt-4 p-4 border border-slate-200 rounded-lg bg-white shadow-sm">
        <h4 className="font-bold text-sm text-slate-700 capitalize mb-3">{phase} Parameters</h4>
        <div className="grid grid-cols-5 gap-3">
          {paramKeys.map((key, index) => (
            <div key={key}>
              <label className="block text-xs font-bold text-slate-500 mb-1 truncate" title={config.parameters[key].label.replace('\n', ' ')}>
                {config.parameters[key].label.split('\n')[0]}
              </label>
              <input
                type="number"
                step="any"
                className="w-full text-sm border border-slate-300 rounded p-2 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none"
                value={gradesData[grade][phase][key] || ''}
                onChange={(e) => handleParamChange(grade, phase, key, e.target.value)}
                onPaste={index === 0 ? (e) => handlePaste(e, grade, phase) : undefined}
                placeholder={index === 0 ? "Paste row here" : ""}
              />
            </div>
          ))}
        </div>
      </div>
    );
  };

  return (
    <div className="bg-slate-50 border border-slate-200 rounded-xl p-6 mb-8 max-w-4xl mx-auto shadow-sm">
      <div className="flex justify-between items-center mb-6">
        <h2 className="text-2xl font-bold text-indigo-950">Manual Entry Form</h2>
        <button type="button" onClick={onCancel} className="text-slate-500 hover:text-slate-700 font-semibold">Cancel</button>
      </div>

      <form onSubmit={handleSubmit}>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
          <div>
            <label className="block text-sm font-bold text-slate-700 mb-1">District</label>
            <input type="text" className="w-full border border-slate-300 rounded p-2 outline-none focus:border-indigo-500" value={district} onChange={e => setDistrict(e.target.value)} />
          </div>
          <div>
            <label className="block text-sm font-bold text-slate-700 mb-1">Mandal</label>
            <input type="text" className="w-full border border-slate-300 rounded p-2 outline-none focus:border-indigo-500" value={mandal} onChange={e => setMandal(e.target.value)} />
          </div>
          <div>
            <label className="block text-sm font-bold text-slate-700 mb-1">School Name <span className="text-red-500">*</span></label>
            <input type="text" className="w-full border border-slate-300 rounded p-2 outline-none focus:border-indigo-500" value={schoolName} onChange={e => setSchoolName(e.target.value)} required />
          </div>
        </div>

        <div className="mb-8">
          <label className="block text-sm font-bold text-slate-700 mb-2">Target Report Phase</label>
          <div className="flex space-x-4">
            {(['Baseline', 'Midline', 'Endline'] as TestPhase[]).map(p => (
              <label key={p} className="flex items-center space-x-2 cursor-pointer">
                <input type="radio" className="text-indigo-600 focus:ring-indigo-500" checked={targetPhase === p} onChange={() => setTargetPhase(p)} />
                <span className="font-semibold text-slate-700">{p}</span>
              </label>
            ))}
          </div>
        </div>

        <div className="space-y-8">
          {[3, 4, 5].map(grade => (
            <div key={grade} className="bg-slate-100 p-5 rounded-xl border border-slate-200">
              <div className="flex items-center justify-between mb-2">
                <h3 className="text-lg font-bold text-slate-800">Grade {grade}</h3>
                <div className="flex items-center space-x-2">
                  <label className="text-sm font-bold text-slate-600">Total Assessed:</label>
                  <input type="number" className="w-20 border border-slate-300 rounded p-1 outline-none focus:border-indigo-500" value={gradesData[grade].assessedCount || ''} onChange={e => handleCountChange(grade, e.target.value)} />
                </div>
              </div>
              
              {renderPhaseInputs(grade, 'baseline')}
              {(targetPhase === 'Midline' || targetPhase === 'Endline') && renderPhaseInputs(grade, 'midline')}
              {targetPhase === 'Endline' && renderPhaseInputs(grade, 'endline')}
            </div>
          ))}
        </div>

        <div className="mt-8 flex justify-end">
          <button type="submit" className="bg-indigo-600 text-white font-bold py-3 px-8 rounded-lg hover:bg-indigo-700 transition-colors shadow-sm">
            Generate Report
          </button>
        </div>
      </form>
    </div>
  );
}
