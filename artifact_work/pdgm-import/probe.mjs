import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const w=Workbook.create();const s=w.worksheets.add('Probe');
s.getRange('A1').setNumberFormat('@');s.getRange('A1').values=[['2026-09-21T00:30:00+03:00']];
s.getRange('A2').values=[["'2026-09-21T00:30:00+03:00"]];
s.getRange('A3').values=[[new Date('2026-09-21T12:00:00Z')]];s.getRange('A3').setNumberFormat('yyyy-mm-dd');
s.getRange('A4').formulas=[['="2026-09-21T00:30:00+03:00"']];
await (await SpreadsheetFile.exportXlsx(w)).save('artifact_work/pdgm-import/probe.xlsx');
