import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const root = process.cwd();
const dir = path.join(root, 'tmp/kanban_update_2026_10_14');
const target = path.join(root, '1. FLOODSENSE-MAIN/KANBAN(WORK-IN-TO-DO-INPROGRESS-DONE)/FloodSense_Kanban_Board.xlsm');
const source = 'C:/Users/DyyYah/.codex/attachments/a8eaa0d-8c33-41f2-946d-3de8f7e66b29/Pasted text.txt'.replace('a8eaa0d', 'a8e4aa0d');
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(target));
const sheet = wb.worksheets.getItemAt(0);
if (process.argv.includes('--inspect')) {
  console.log((await wb.inspect({kind:'workbook,sheet,table',maxChars:3000})).ndjson);
  const img = await wb.render({sheetName:sheet.name,range:'A1:E4',scale:2,format:'png'});
  await fs.writeFile(path.join(dir,'before.png'),new Uint8Array(await img.arrayBuffer()));
  process.exit(0);
}
const text = (await fs.readFile(source,'utf8')).replace(/^\uFEFF/, '');
let category = '';
const items = [];
for (const line of text.split(/\r?\n/)) {
  if (!line.trim()) continue;
  if (!line.startsWith('- ')) { category=line.trim(); continue; }
  const match = line.match(/^- (.+?) — (.+)$/);
  if (!match) throw new Error(`Unparsed source line: ${line}`);
  items.push({category,status:match[1],text:match[2],original:line.slice(2)});
}
const lanes = {A:[],B:[],C:[]};
const excluded = [];
for (const item of items) {
  if (item.status === 'Not done by scope') excluded.push(item);
  else if (item.status.startsWith('Done')) lanes.C.push(item);
  else if (item.status === 'In progress') lanes.B.push(item);
  else if (['Not done','Deferred'].includes(item.status)) lanes.A.push(item);
  else throw new Error(`Unsupported status: ${item.status}`);
}
const maxCount = Math.max(...Object.values(lanes).map(a=>a.length));
sheet.getRange(`A2:D${maxCount+1}`).format.columnWidth=62;
sheet.getRange(`A2:D${maxCount+1}`).format.font={name:'Calibri',size:11};
sheet.getRange(`A2:D${maxCount+1}`).format.wrapText=true;
sheet.getRange(`A2:D${maxCount+1}`).format.verticalAlignment='top';
const written=[];
for (const [col,list] of Object.entries(lanes)) {
  if (!list.length) continue;
  const cards=list.map(item=>`${item.category}\n${item.original}`);
  sheet.getRange(`${col}2:${col}${list.length+1}`).values=cards.map(card=>[card]);
  written.push(...list.map((item,i)=>({...item,cell:`${col}${i+2}`,value:cards[i]})));
}
// Fit all descriptions, keeping all supplied wording and categories visible.
for (let i=0;i<maxCount;i++) {
  const lineCount=Math.max(...Object.values(lanes).map(list=>{
    const item=list[i];
    if (!item) return 1;
    return Math.ceil(item.category.length/58)+Math.ceil(item.original.length/58)+1;
  }));
  sheet.getRange(`A${i+2}:D${i+2}`).format.rowHeight=Math.min(409,Math.max(60,lineCount*15+12));
}
const noteRow=maxCount+3;
for (const [i,item] of excluded.entries()) {
  const row=noteRow+i;
  sheet.mergeCells(`A${row}:D${row}`);
  const value=`OUT OF SCOPE (not a pending task)\n${item.category}\n${item.original}`;
  sheet.getRange(`A${row}`).values=[[value]];
  sheet.getRange(`A${row}:D${row}`).format={font:{name:'Calibri',size:11},wrapText:true,verticalAlignment:'top',rowHeight:85};
  written.push({...item,cell:`A${row}`,value});
}
wb.recalculate();
for (const item of written) {
  if (sheet.getRange(item.cell).values[0][0]!==item.value) throw new Error(`Content mismatch ${item.cell}`);
}
console.log(JSON.stringify({total:items.length,todo:lanes.A.length,inProgress:lanes.B.length,done:lanes.C.length,partiallyDone:0,outOfScope:excluded.length}));
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!',options:{useRegex:true,maxResults:20},maxChars:1500})).ndjson);
await fs.writeFile(path.join(dir,'manifest.json'),JSON.stringify({target,written,maxCount,noteRow},null,2));
for (const [name,range] of [['after_top','A1:D4'],['after_bottom',`A${maxCount}:D${noteRow}`]]) {
  const img=await wb.render({sheetName:sheet.name,range,scale:1.3,format:'png'});
  await fs.writeFile(path.join(dir,`${name}.png`),new Uint8Array(await img.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(wb)).save(path.join(dir,'staging.xlsx'));
