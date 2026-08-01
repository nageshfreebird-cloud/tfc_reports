import React, { useState } from 'react';
import { read, utils } from 'xlsx';
import { SchoolData, TestPhase, ReportConfig } from './types';
import SchoolReportPDF from './components/SchoolReportPDF';
import ManualEntryForm from './components/ManualEntryForm';
import html2canvas from 'html2canvas';
import { generateWordReport } from './utils/wordExport';
import { FileText, Download, Upload, Settings } from 'lucide-react';
import Logo from './components/Logo';

const DEFAULT_CONFIG: ReportConfig = {
  parameters: {
    KNOW: { label: 'KNOW', max: 10 },
    READ: { label: 'READ', max: 8 },
    SPELL: { label: 'SPELL', max: 8 },
    CWR: { label: 'CAMERA WORD\nREAD', max: 12 },
    CWS: { label: 'CAMERA WORD\nSPELL', max: 12 }
  },
  phases: {
    Baseline: {
      aboutText: "The Baseline Assessment is conducted in the beginning of the smart classroom program of Teach For Change. In this assessment, the students are individually assessed in 5 parameters of English Language Literacy.",
      nextStepsTitle: "Next Steps",
      nextStepsBullets: ["", "", "", ""]
    },
    Midline: {
      aboutText: "The Midline Assessment is conducted in the mid of year of the smart classroom program of Teach For Change. In this assessment, the students are individually assessed in 5 parameters of English Language Literacy.",
      nextStepsTitle: "Next Steps",
      nextStepsBullets: [
        "Focus on phonics and small group help.",
        "Make classrooms rich with reading books.",
        "Train teachers on effective ESL methods.",
        "Set clear goals and assess progress often"
      ]
    },
    Endline: {
      aboutText: "The Endline Assessment is conducted at the end of the year of the smart classroom program of Teach For Change. In this assessment, the students are individually assessed in 5 parameters of English Language Literacy.",
      nextStepsTitle: "THANK YOU",
      nextStepsBullets: [
        "Thank you very much for your valuable support throughout this year. We truly appreciate your collaboration and guidance. As we move forward, we look forward to working together in the coming year to further enhance English language skills among students in Grades 3, 4, and 5."
      ]
    }
  }
};

export default function App() {
  const [schools, setSchools] = useState<SchoolData[]>([]);
  const [reportSchool, setReportSchool] = useState<string>('');
  const [reportPhase, setReportPhase] = useState<TestPhase>('Baseline');
  const [isExportingWord, setIsExportingWord] = useState(false);
  const [isPrintingNative, setIsPrintingNative] = useState(false);
  const [config, setConfig] = useState<ReportConfig>(DEFAULT_CONFIG);
  const [showSettings, setShowSettings] = useState(false);
  const [sheetUrl, setSheetUrl] = useState('');
  const [isLoadingSheet, setIsLoadingSheet] = useState(false);
  const [isManualMode, setIsManualMode] = useState(false);

  // Core parsing logic separated so it can be used by both file and URL
  const parseWorkbook = (workbook: any) => {
    try {
      // Look for a sheet named "Reports", otherwise fallback to the first sheet
      const targetSheetName = workbook.SheetNames.find((name: string) => name.toLowerCase() === 'reports');
      const sheetName = targetSheetName || workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      // Get rows as arrays
      const rows: any[][] = utils.sheet_to_json(worksheet, { header: 1 });
        
        const parsedSchools = new Map<string, SchoolData>();

        // Skip headers (start from row 2 or 3)
        // Find the first row that actually has a district
        let startRow = 2;
        for (let i = 0; i < rows.length; i++) {
          if (rows[i][0] && rows[i][0].toString().toLowerCase() !== 'district') {
            startRow = i;
            break;
          }
        }

        let lastDistrict = '';
        let lastMandal = '';
        let lastSchool = '';

        for (let i = startRow; i < rows.length; i++) {
          const r = rows[i];
          if (!r || r.length < 5) continue; // Skip empty rows

          if (r[0]) lastDistrict = r[0].toString().trim();
          if (r[1]) lastMandal = r[1].toString().trim();
          if (r[2]) lastSchool = r[2].toString().trim();

          const district = lastDistrict;
          const mandal = lastMandal;
          const schoolName = lastSchool;
          const classNameStr = (r[3] || '').toString().toLowerCase();

          // If we still don't have a school name, skip
          if (!schoolName) continue;

          let grade: 3 | 4 | 5 = 3;
          if (classNameStr.includes('4')) grade = 4;
          else if (classNameStr.includes('5')) grade = 5;

          const strength = Number(r[4]) || 0;

          if (!parsedSchools.has(schoolName)) {
            parsedSchools.set(schoolName, {
              District: district,
              Mandal: mandal,
              School_Name: schoolName,
              classes: {}
            });
          }

          const school = parsedSchools.get(schoolName)!;

          school.classes[grade] = {
            assessedCount: strength,
            baseline: {
              KNOW: Number(r[5]) || 0,
              READ: Number(r[6]) || 0,
              SPELL: Number(r[7]) || 0,
              CWR: Number(r[8]) || 0,
              CWS: Number(r[9]) || 0,
              Total: Number(r[10]) || 0,
            },
            midline: {
              KNOW: Number(r[12]) || 0,
              READ: Number(r[13]) || 0,
              SPELL: Number(r[14]) || 0,
              CWR: Number(r[15]) || 0,
              CWS: Number(r[16]) || 0,
              Total: Number(r[17]) || 0,
            },
            endline: {
              KNOW: Number(r[19]) || 0,
              READ: Number(r[20]) || 0,
              SPELL: Number(r[21]) || 0,
              CWR: Number(r[22]) || 0,
              CWS: Number(r[23]) || 0,
              Total: Number(r[24]) || 0,
            }
          };
        }

        const schoolsList = Array.from(parsedSchools.values());
        setSchools(schoolsList);
        if (schoolsList.length > 0) {
          setReportSchool(schoolsList[0].School_Name);
        }
        return schoolsList;
      } catch (err) {
        console.error(err);
        alert('Failed to parse spreadsheet. Make sure it matches the expected format.');
        return [];
      }
  };

  // Parse Google Sheet Exported CSV or XLSX from File
  const processSpreadsheetFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const fileData = e.target?.result;
      if (fileData) {
        const workbook = read(fileData, { type: 'binary' });
        parseWorkbook(workbook);
        alert('Successfully imported data from file!');
      }
    };
    reader.readAsBinaryString(file);
  };

  // Parse Google Sheet from URL
  const processSpreadsheetUrl = async (url: string) => {
    if (!url.trim()) return;
    setIsLoadingSheet(true);
    try {
      // Extract the ID from standard google sheet links
      // e.g. https://docs.google.com/spreadsheets/d/1wzd2oEq_0WR0psBFhVb_vQ9WiHL9XdnQJdTqMWY96o0/edit
      const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) {
        throw new Error('Invalid Google Sheets URL. Could not find Document ID.');
      }
      
      const docId = match[1];
      // Export as xlsx so we get all sheets, not just the first one
      const xlsxUrl = `https://docs.google.com/spreadsheets/d/${docId}/export?format=xlsx`;
      
      // We use allorigins as a reliable free CORS proxy to avoid browser blocks
      const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(xlsxUrl)}`;
      
      const response = await fetch(proxyUrl);
      if (!response.ok) throw new Error('Failed to fetch from Google Sheets.');
      
      const arrayBuffer = await response.arrayBuffer();
      
      // Parse the XLSX ArrayBuffer using XLSX
      const workbook = read(arrayBuffer, { type: 'array' });
      const schoolsList = parseWorkbook(workbook);
      
      if (schoolsList && schoolsList.length > 0) {
        alert('Successfully imported data from Google Link!');
        setSheetUrl('');
      }
    } catch (err: any) {
      console.error(err);
      alert(`Error fetching from link: ${err.message}`);
    } finally {
      setIsLoadingSheet(false);
    }
  };

  const handleConfigChange = (key: keyof ReportConfig['parameters'], field: 'label' | 'max', value: string | number) => {
    setConfig(prev => ({
      ...prev,
      parameters: {
        ...prev.parameters,
        [key]: {
          ...prev.parameters[key],
          [field]: value
        }
      }
    }));
  };

  if ((isPrintingNative || isExportingWord) && reportSchool) {
    return (
      <div className="bg-white m-0 p-0 w-full h-full min-h-screen flex justify-center items-start">
        <style>
          {`
            @page { size: A4 portrait; margin: 0; }
            body { margin: 0; padding: 0; background: white; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          `}
        </style>
        <SchoolReportPDF
          school={schools.find(s => s.School_Name === reportSchool)!}
          phase={reportPhase}
          config={config}
        />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans">
      <header className="bg-indigo-950 text-white shadow-xl relative overflow-hidden">
        <div className="absolute inset-0 bg-[radial-gradient(circle_at_top_right,rgba(99,102,241,0.2),transparent_50%)]" />
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4.5 flex items-center justify-between relative z-10">
          <div className="flex items-center space-x-3.5">
            <div className="w-11 h-11 flex items-center justify-center bg-white/10 backdrop-blur-md rounded-xl shadow-lg p-0.5 border border-white/20">
              <Logo className="w-10 h-10" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h1 className="text-lg font-black tracking-tight text-white">Report Generator</h1>
              </div>
              <p className="text-xs text-slate-400 font-medium mt-0.5">Standalone PDF & Word Generation Tool</p>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 w-full max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
        
        {isManualMode ? (
          <ManualEntryForm 
            config={config} 
            onSubmit={(data, phase) => {
              setSchools([data]);
              setReportSchool(data.School_Name);
              setReportPhase(phase);
              setIsManualMode(false);
            }} 
            onCancel={() => setIsManualMode(false)} 
          />
        ) : (
          <>
            {/* Step 1: Upload */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-xl font-bold flex items-center"><Upload className="w-5 h-5 mr-2 text-indigo-600"/> Step 1: Import Data</h2>
                <button onClick={() => setIsManualMode(true)} className="bg-indigo-100 text-indigo-700 hover:bg-indigo-200 font-bold py-2 px-4 rounded-lg text-sm transition-colors">
                  Manual Entry Mode
                </button>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="flex flex-col justify-center border-2 border-dashed border-slate-300 rounded-xl p-6 bg-slate-50 hover:bg-slate-100 transition-colors h-full">
              <h3 className="font-bold text-slate-700 mb-2 text-sm text-center">Upload File</h3>
              <input
                type="file"
                accept=".xlsx,.csv"
                onChange={(e) => {
                  if (e.target.files?.[0]) processSpreadsheetFile(e.target.files[0]);
                }}
                className="mb-4 mx-auto block text-sm w-full max-w-[200px]"
              />
              <p className="text-[11px] text-slate-500 text-center leading-relaxed">
                Upload your Excel or CSV file.
              </p>
            </div>

            <div className="flex flex-col justify-center border-2 border-slate-200 rounded-xl p-6 bg-white h-full">
              <h3 className="font-bold text-slate-700 mb-2 text-sm text-center">Or Paste Google Sheets Link</h3>
              <div className="flex space-x-2">
                <input
                  type="url"
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  value={sheetUrl}
                  onChange={(e) => setSheetUrl(e.target.value)}
                  className="flex-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs focus:outline-none focus:border-indigo-500"
                />
                <button
                  onClick={() => processSpreadsheetUrl(sheetUrl)}
                  disabled={!sheetUrl || isLoadingSheet}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-bold disabled:opacity-50 transition-colors"
                >
                  {isLoadingSheet ? 'Loading...' : 'Import'}
                </button>
              </div>
              <p className="text-[11px] text-slate-500 text-center mt-3 leading-relaxed">
                Paste the sharing link to your Google Sheet. It must be accessible to "Anyone with the link".
              </p>
            </div>
          </div>
        </div>

        {schools.length > 0 && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Step 2: Configure Parameters */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-bold flex items-center"><Settings className="w-5 h-5 mr-2 text-indigo-600"/> Settings (Optional)</h2>
                <button onClick={() => setShowSettings(!showSettings)} className="text-sm text-indigo-600 font-bold hover:underline">
                  {showSettings ? 'Hide' : 'Show'} Parameters
                </button>
              </div>
              
              {showSettings && (
                <div className="space-y-6 max-h-[60vh] overflow-y-auto pr-2">
                  <div>
                    <h3 className="font-bold text-sm text-slate-800 mb-2">Parameters</h3>
                    <div className="space-y-3">
                      {(Object.keys(config.parameters) as Array<keyof ReportConfig['parameters']>).map(key => (
                        <div key={key} className="flex space-x-4 items-center bg-slate-50 p-3 rounded-lg border border-slate-100">
                          <div className="w-20 font-bold text-sm text-slate-700">{key}</div>
                          <div className="flex-1">
                            <label className="text-[10px] text-slate-500 uppercase block mb-1">Label</label>
                            <input 
                              type="text" 
                              value={config.parameters[key].label} 
                              onChange={(e) => handleConfigChange(key, 'label', e.target.value)}
                              className="w-full text-sm border border-slate-200 rounded p-1"
                            />
                          </div>
                          <div className="w-24">
                            <label className="text-[10px] text-slate-500 uppercase block mb-1">Max Marks</label>
                            <input 
                              type="number" 
                              value={config.parameters[key].max} 
                              onChange={(e) => handleConfigChange(key, 'max', Number(e.target.value))}
                              className="w-full text-sm border border-slate-200 rounded p-1"
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div>
                    <h3 className="font-bold text-sm text-slate-800 mb-2">Phase Texts</h3>
                    <div className="space-y-4">
                      {(['Baseline', 'Midline', 'Endline'] as TestPhase[]).map(phase => (
                        <div key={phase} className="bg-slate-50 p-4 rounded-lg border border-slate-100 space-y-3">
                          <h4 className="font-bold text-sm text-indigo-600">{phase}</h4>
                          <div>
                            <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">About Text</label>
                            <textarea 
                              value={config.phases[phase].aboutText}
                              onChange={(e) => {
                                setConfig(prev => ({
                                  ...prev,
                                  phases: {
                                    ...prev.phases,
                                    [phase]: { ...prev.phases[phase], aboutText: e.target.value }
                                  }
                                }));
                              }}
                              className="w-full text-xs border border-slate-200 rounded p-2 min-h-[60px]"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Next Steps Title</label>
                            <input 
                              type="text"
                              value={config.phases[phase].nextStepsTitle}
                              onChange={(e) => {
                                setConfig(prev => ({
                                  ...prev,
                                  phases: {
                                    ...prev.phases,
                                    [phase]: { ...prev.phases[phase], nextStepsTitle: e.target.value }
                                  }
                                }));
                              }}
                              className="w-full text-xs border border-slate-200 rounded p-2"
                            />
                          </div>
                          <div>
                            <label className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Next Steps / Content (One per line)</label>
                            <textarea 
                              value={config.phases[phase].nextStepsBullets.join('\n')}
                              onChange={(e) => {
                                setConfig(prev => ({
                                  ...prev,
                                  phases: {
                                    ...prev.phases,
                                    [phase]: { ...prev.phases[phase], nextStepsBullets: e.target.value.split('\n') }
                                  }
                                }));
                              }}
                              className="w-full text-xs border border-slate-200 rounded p-2 min-h-[80px]"
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Step 3: Print */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col justify-between">
              <div>
                <h2 className="text-xl font-bold mb-4 flex items-center"><Download className="w-5 h-5 mr-2 text-indigo-600"/> Step 2: Generate Report</h2>
                <div className="space-y-4">
                  <div>
                    <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Select School</label>
                    <select
                      value={reportSchool}
                      onChange={(e) => setReportSchool(e.target.value)}
                      className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-800"
                    >
                      {schools.map(s => (
                        <option key={s.School_Name} value={s.School_Name}>{s.School_Name} ({s.District})</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-xs font-bold text-slate-500 uppercase block mb-1">Select Phase</label>
                    <select
                      value={reportPhase}
                      onChange={(e) => setReportPhase(e.target.value as TestPhase)}
                      className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-800"
                    >
                      <option value="Baseline">Baseline</option>
                      <option value="Midline">Midline</option>
                      <option value="Endline">Endline</option>
                    </select>
                  </div>
                </div>
              </div>

              <div className="space-y-3 mt-6">
                <button
                  onClick={() => {
                    setIsPrintingNative(true);
                    setTimeout(() => {
                      window.print();
                      setIsPrintingNative(false);
                    }, 500);
                  }}
                  disabled={!reportSchool || isPrintingNative}
                  className="w-full py-4 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white rounded-xl text-sm font-black shadow-lg disabled:opacity-50 active:scale-[0.98] transition-all flex items-center justify-center space-x-2"
                >
                  <Download className="w-5 h-5" />
                  <span>Download PDF Report</span>
                </button>

                <button
                  onClick={() => {
                    setIsExportingWord(true);
                    setTimeout(async () => {
                      try {
                        const getChartDataUrl = async (id: string) => {
                          const el = document.getElementById(id);
                          if (!el) return '';
                          const canvas = await html2canvas(el, { scale: 2 });
                          return canvas.toDataURL('image/png');
                        };
                        
                        const [chart3, chart4, chart5] = await Promise.all([
                          getChartDataUrl('chart-grade-3'),
                          getChartDataUrl('chart-grade-4'),
                          getChartDataUrl('chart-grade-5')
                        ]);
                        
                        const selectedSchool = schools.find(s => s.School_Name === reportSchool)!;
                        const blob = await generateWordReport(selectedSchool, reportPhase, config, {
                          grade3: chart3,
                          grade4: chart4,
                          grade5: chart5
                        });
                        
                        const url = URL.createObjectURL(blob);
                        const link = document.createElement('a');
                        link.href = url;
                        link.download = `${selectedSchool.School_Name.replace(/\s+/g, '_')}_${reportPhase}_Report.docx`;
                        link.click();
                        URL.revokeObjectURL(url);
                      } catch (err) {
                        console.error('Error exporting word:', err);
                        alert('Failed to generate Word document.');
                      } finally {
                        setIsExportingWord(false);
                      }
                    }, 500);
                  }}
                  disabled={!reportSchool || isExportingWord || isPrintingNative}
                  className="w-full py-4 bg-gradient-to-r from-blue-600 to-cyan-600 hover:from-blue-500 hover:to-cyan-500 text-white rounded-xl text-sm font-black shadow-lg disabled:opacity-50 active:scale-[0.98] transition-all flex items-center justify-center space-x-2"
                >
                  {isExportingWord ? (
                    <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  ) : (
                    <FileText className="w-5 h-5" />
                  )}
                  <span>{isExportingWord ? 'Preparing Word...' : 'Download Word Report'}</span>
                </button>
              </div>
            </div>
          </div>
        )}
          </>
        )}
      </main>
    </div>
  );
}
