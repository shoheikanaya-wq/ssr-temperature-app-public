const VERSION="1.3";
const input=document.getElementById("input"),result=document.getElementById("result");
const wordFile=document.getElementById("wordFile");

async function readDocx(file){
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const doc=zip.file("word/document.xml");
  if(!doc)throw new Error("Word文書の本文を読み取れませんでした。");
  const xml=await doc.async("string");
  const parser=new DOMParser();
  const dom=parser.parseFromString(xml,"application/xml");
  if(dom.getElementsByTagName("parsererror").length)throw new Error("Word文書の解析に失敗しました。");
  const paragraphs=[...dom.getElementsByTagNameNS("*","p")];
  const lines=paragraphs.map(p=>{
    let line="";
    for(const node of [...p.getElementsByTagNameNS("*","t")]) line+=node.textContent||"";
    return line.trimEnd();
  });
  return lines.join("\n").replace(/\n{3,}/g,"\n\n").trim();
}

if(wordFile){
  wordFile.addEventListener("change",async()=>{
    const file=wordFile.files?.[0];
    if(!file)return;
    try{
      result.textContent="Wordを読み込み中…";
      const text=await readDocx(file);
      if(!text)throw new Error("Word文書に読み取れる文字がありません。");
      input.value=text;
      pendingGenerate=false;
      renderMissingData();
      result.textContent="Word読み込み完了："+file.name;
    }catch(e){
      wordFile.value="";
      result.textContent="Word読込エラー: "+e.message;
    }
  });
}

const esc=s=>s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
function colA(row){return "A"+row}
const MACHINE_LAYOUT={
  "MM18":[2,14,14,14],
  "FE01":[10,8],
  "FE04":[10,12],
  "FM81":[7,7,6],
  "FM83-1":[18,17],
  "FM83-2":[6,6],
  "FM84":[6,14,8],
  "FM96":[6,6,6],
  "FM97":[7,7,4],
  "FM19":[6,6,6,6,6],
  "FM22":[12,9,9],
  "FM01":[7,7,6],
  "FM07":[12,12,8]
};
const EXPECTED_COUNTS=Object.fromEntries(
  Object.entries(MACHINE_LAYOUT).map(([machine,rows])=>[
    machine,
    2+rows.reduce((sum,n)=>sum+n,0)
  ])
);
const MACHINE_RE=/(FM\s*83-1|FM\s*83-2|FM\s*81|FM\s*84|FM\s*96|FM\s*97|FM\s*19|FM\s*22|FM\s*01|FM\s*07|FE\s*01|FE\s*04|MM\s*18)/i;
function normalizeMachineName(s){return String(s||"").toUpperCase().replace(/\s+/g,"")}

const TEMPLATE_MAP={
  "MM18":{panel:["D2","F2"],ssr:["B4","C4","B6","C6","D6","E6","F6","G6","H6","I6","J6","K6","L6","M6","N6","O6","B8","C8","D8","E8","F8","G8","H8","I8","J8","K8","L8","M8","N8","O8","B10","C10","D10","E10","F10","G10","H10","I10","J10","K10","L10","M10","N10","O10"]},
  "FE01":{panel:["D12","F12"],ssr:["B14","C14","D14","E14","F14","G14","H14","I14","J14","K14","B16","C16","F16","G16","H16","I16","J16","K16"]},
  "FE04":{panel:["D18","F18"],ssr:["B20","C20","D20","E20","F20","G20","H20","I20","J20","K20","B22","C22","D22","E22","F22","G22","H22","I22","J22","K22","L22","M22"]},
  "FM81":{panel:["D24","F24"],ssr:["B26","C26","D26","E26","F26","G26","H26","B28","C28","D28","E28","F28","G28","H28","B30","C30","D30","E30","F30","G30"]},
  "FM83-1":{panel:["D32","F32"],ssr:["B34","C34","D34","E34","F34","G34","H34","I34","J34","K34","L34","M34","N34","O34","P34","Q34","R34","S34","B36","C36","D36","E36","F36","G36","H36","I36","J36","K36","L36","M36","N36","O36","P36","Q36","R36"]},
  "FM83-2":{panel:["D38","F38"],ssr:["B40","C40","D40","E40","F40","G40","B42","C42","D42","E42","F42","G42"]},
  "FM84":{panel:["D45","F45"],ssr:["B47","C47","D47","E47","F47","G47","B49","C49","D49","E49","F49","G49","H49","I49","J49","K49","L49","M49","N49","O49","B51","C51","D51","E51","F51","G51","H51","I51"]},
  "FM96":{panel:["D54","F54"],ssr:["B56","C56","D56","E56","F56","G56","B58","C58","D58","E58","F58","G58","B60","C60","D60","E60","F60","G60"]},
  "FM97":{panel:["M54","O54"],ssr:["K56","L56","M56","N56","O56","P56","Q56","K58","L58","M58","N58","O58","P58","Q58","K60","L60","M60","N60"]},
  "FM19":{panel:["D63","F63"],ssr:["B65","C65","D65","E65","F65","G65","B67","C67","D67","E67","F67","G67","B69","C69","D69","E69","F69","G69","B71","C71","D71","E71","F71","G71","B73","C73","D73","E73","F73","G73"]},
  "FM22":{panel:["M63","O63"],ssr:["K65","L65","M65","N65","O65","P65","Q65","R65","S65","T65","U65","V65","K67","L67","M67","N67","O67","P67","Q67","R67","S67","K69","L69","M69","N69","O69","P69","Q69","R69","S69"]},
  "FM01":{panel:["D76","F76"],ssr:["B78","C78","D78","E78","F78","G78","H78","B80","C80","D80","E80","F80","G80","H80","B82","C82","D82","E82","F82","G82"]},
  "FM07":{panel:["M76","O76"],ssr:["K78","L78","M78","N78","O78","P78","Q78","R78","S78","T78","U78","V78","K80","L80","M80","N80","O80","P80","Q80","R80","S80","T80","U80","V80","O82","P82","Q82","R82","S82","T82","U82","V82"]}
};
const VERTICAL_COL={"MM18":"B","FE01":"C","FE04":"D","FM81":"E","FM83-1":"F","FM83-2":"G","FM84":"H","FM96":"I","FM97":"J","FM19":"K","FM22":"L","FM01":"M","FM07":"N"};

function parseMeasurementValues(line){
  const chunks=line.trim().split(/\s+/).filter(Boolean);
  const values=[];
  let invalid=false;

  for(const chunk of chunks){
    if(/^-?\d{1,3}\.\d$/.test(chunk)){
      values.push(chunk);
      continue;
    }

    if(/^\d{2}\.\d(?:\d{2}\.\d)+$/.test(chunk)){
      const joined=chunk.match(/\d{2}\.\d/g)||[];
      values.push(...joined);
      continue;
    }

    if(/\d/.test(chunk)){
      values.push(chunk);
      invalid=true;
    }
  }

  return {values,invalid};
}

function countNumbers(text){
  return parseMeasurementValues(text).values.length;
}

function parseMachineBlocks(text){
  const blocks=[];
  let current=null;
  for(const raw of text.replace(/\r/g,"").split("\n")){
    const hit=raw.match(MACHINE_RE);
    if(hit){
      if(current)blocks.push(current);
      const machine=normalizeMachineName(hit[1]);
      const after=raw.replace(hit[0]," ").trim();
      current={machine,lines:[],text:"",count:0};
      if(after){
        current.lines.push(after);
        current.text=after;
        current.count+=countNumbers(after);
      }
    }else if(current){
      current.lines.push(raw);
      current.text+=(current.text?"\n":"")+raw;
      current.count+=countNumbers(raw);
    }
  }
  if(current)blocks.push(current);
  return blocks;
}

function parseDateFromText(text){
  let m=text.match(/(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})\s*日?/);
  if(!m)m=text.match(/(20\d{2})(\d{2})(\d{2})/);
  if(!m)return null;
  const y=Number(m[1]),mo=Number(m[2]),d=Number(m[3]);
  return {y,mo,d,full:y+"年"+mo+"月"+d+"日",short:mo+"月"+d+"日",file:String(y)+String(mo).padStart(2,"0")+String(d).padStart(2,"0")};
}

function parseFactoryTemps(text){
  const out={1:null,2:null,3:null};
  for(const line of text.replace(/\r/g,"").split("\n")){
    let idx=null;
    if(/①|まるいち|丸一|丸1|構内.*1/i.test(line))idx=1;
    else if(/②|まるに|丸二|丸2|構内.*2/i.test(line))idx=2;
    else if(/③|まるさん|丸三|丸3|構内.*3/i.test(line))idx=3;
    if(!idx)continue;
    const vals=parseMeasurementValues(line).values.filter(v=>/^\d{1,3}\.\d$/.test(v));
    if(vals.length)out[idx]=Number(vals[vals.length-1]);
  }
  return out;
}

function getSheetTarget(wbdoc,reldoc,name){
  const sheet=[...wbdoc.getElementsByTagName("sheet")].find(x=>x.getAttribute("name")===name);
  if(!sheet)return null;
  const rid=sheet.getAttribute("r:id");
  const rel=[...reldoc.getElementsByTagName("Relationship")].find(x=>x.getAttribute("Id")===rid);
  if(!rel)return null;
  let target=rel.getAttribute("Target").replace(/^\//,"");
  if(!target.startsWith("xl/"))target="xl/"+target.replace(/^\.\//,"");
  return target;
}

function ensureRow(doc,rowNum){
  const ns=doc.documentElement.namespaceURI;
  const sheetData=doc.getElementsByTagName("sheetData")[0];
  let row=[...sheetData.getElementsByTagName("row")].find(r=>Number(r.getAttribute("r"))===rowNum);
  if(!row){
    row=doc.createElementNS(ns,"row");
    row.setAttribute("r",String(rowNum));
    let inserted=false;
    for(const existing of [...sheetData.getElementsByTagName("row")]){
      if(Number(existing.getAttribute("r"))>rowNum){
        sheetData.insertBefore(row,existing);
        inserted=true;
        break;
      }
    }
    if(!inserted)sheetData.appendChild(row);
  }
  return row;
}

function setCellNumber(doc,ref,value){
  const ns=doc.documentElement.namespaceURI;
  const rowNum=Number((ref.match(/\d+/)||["0"])[0]);
  const row=ensureRow(doc,rowNum);
  let cell=[...row.getElementsByTagName("c")].find(c=>c.getAttribute("r")===ref);
  if(!cell){
    cell=doc.createElementNS(ns,"c");
    cell.setAttribute("r",ref);
    row.appendChild(cell);
  }
  cell.removeAttribute("t");
  while(cell.firstChild)cell.removeChild(cell.firstChild);
  const v=doc.createElementNS(ns,"v");
  v.textContent=String(Number(value));
  cell.appendChild(v);
}

function setCellText(doc,ref,value){
  const ns=doc.documentElement.namespaceURI;
  const rowNum=Number((ref.match(/\d+/)||["0"])[0]);
  const row=ensureRow(doc,rowNum);
  let cell=[...row.getElementsByTagName("c")].find(c=>c.getAttribute("r")===ref);
  if(!cell){
    cell=doc.createElementNS(ns,"c");
    cell.setAttribute("r",ref);
    row.appendChild(cell);
  }
  cell.setAttribute("t","inlineStr");
  while(cell.firstChild)cell.removeChild(cell.firstChild);
  const is=doc.createElementNS(ns,"is");
  const t=doc.createElementNS(ns,"t");
  t.setAttribute("xml:space","preserve");
  t.textContent=String(value??"");
  is.appendChild(t);
  cell.appendChild(is);
}

function machineValues(machine){
  const range=machineBlockRange(machine);
  if(!range)return [];
  const rawLines=range.lines.slice(range.start,range.end);
  if(rawLines.length)rawLines[0]=rawLines[0].replace(MACHINE_RE," ").trim();
  const values=[];
  for(const line of rawLines){
    values.push(...parseMeasurementValues(line).values.filter(v=>/^-?\d{1,3}\.\d$/.test(v)).map(Number));
  }
  return values;
}

async function build(){
 try{
  const text=input.value.replace(/\r/g,"").trim();
  if(!text)throw new Error("Word測定データを選択してください。");
  result.textContent="生成中…";
  if(typeof SSR_TEMPLATE_B64!=="string"||!SSR_TEMPLATE_B64)throw new Error("標準ひな形を読み込めませんでした。");

  const binary=atob(SSR_TEMPLATE_B64);
  const bytes=new Uint8Array(binary.length);
  for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
  const zip=await JSZip.loadAsync(bytes.buffer);

  const parser=new DOMParser(),ser=new XMLSerializer();
  const wbxml=await zip.file("xl/workbook.xml").async("string");
  const wbdoc=parser.parseFromString(wbxml,"application/xml");
  const relxml=await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const reldoc=parser.parseFromString(relxml,"application/xml");

  const mainTarget=getSheetTarget(wbdoc,reldoc,"リレー配置と測定温度記入シート");
  const verticalTarget=getSheetTarget(wbdoc,reldoc,"縦データ");
  if(!mainTarget||!verticalTarget)throw new Error("標準ひな形のシート構成を確認できませんでした。");

  const mainDoc=parser.parseFromString(await zip.file(mainTarget).async("string"),"application/xml");
  const verticalDoc=parser.parseFromString(await zip.file(verticalTarget).async("string"),"application/xml");

  const date=parseDateFromText(text);
  if(date){
    setCellText(mainDoc,"O1",date.full);
    for(const col of Object.values(VERTICAL_COL))setCellText(verticalDoc,col+"2",date.short);
  }

  const factory=parseFactoryTemps(text);
  const factoryText=[1,2,3].map(i=>"①②③"[i-1]+(factory[i]==null?"—":factory[i].toFixed(1)+"℃")).join("　");
  setCellText(mainDoc,"S1",factoryText);

  let measured=0;
  for(const machine of Object.keys(TEMPLATE_MAP)){
    const vals=machineValues(machine);
    if(!vals.length)continue;
    const expected=EXPECTED_COUNTS[machine];
    if(vals.length!==expected)throw new Error(machine+" のデータを修正してから保存してください。");
    measured++;

    const map=TEMPLATE_MAP[machine];
    setCellNumber(mainDoc,map.panel[0],vals[0]);
    setCellNumber(mainDoc,map.panel[1],vals[1]);
    map.ssr.forEach((ref,i)=>setCellNumber(mainDoc,ref,vals[i+2]));

    const vcol=VERTICAL_COL[machine];
    setCellNumber(verticalDoc,vcol+"3",Math.max(vals[0],vals[1]));
    vals.slice(2).forEach((v,i)=>setCellNumber(verticalDoc,vcol+String(6+i),v));
  }

  if(!measured)throw new Error("機台データを認識できませんでした。Wordデータを確認してください。");

  zip.file(mainTarget,ser.serializeToString(mainDoc));
  zip.file(verticalTarget,ser.serializeToString(verticalDoc));

  let calc=wbdoc.getElementsByTagName("calcPr")[0];
  if(!calc){
    calc=wbdoc.createElementNS(wbdoc.documentElement.namespaceURI,"calcPr");
    wbdoc.documentElement.appendChild(calc);
  }
  calc.setAttribute("fullCalcOnLoad","1");
  calc.setAttribute("forceFullCalc","1");
  calc.setAttribute("calcMode","auto");
  zip.file("xl/workbook.xml",ser.serializeToString(wbdoc));

  const out=await zip.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(out);
  a.download=date?date.file+".xlsx":"SSR温測_生成.xlsx";
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),3000);
  result.textContent="生成完了。\n標準ひな形へ "+measured+" 機台を反映しました。";
 }catch(e){
  result.textContent="エラー: "+e.message;
 }
}

let pendingGenerate=false;
document.getElementById("generate").addEventListener("click",()=>{
  const issueCount=renderMissingData();
  if(issueCount>0){
    pendingGenerate=true;
    result.textContent="要修正データを確認してください。";
    missingBox?.scrollIntoView({behavior:"smooth",block:"center"});
    return;
  }
  pendingGenerate=false;
  build();
});
if("serviceWorker" in navigator){
  navigator.serviceWorker.register("./sw.js").then(reg=>{
    window.addEventListener("focus",()=>reg.update());
    document.addEventListener("visibilitychange",()=>{
      if(document.visibilityState==="visible")reg.update();
    });
  }).catch(()=>{});
}

let ssrInstallPrompt=null;
const installBtn=document.getElementById("installApp");
const installGuide=document.getElementById("installGuide");
window.addEventListener("beforeinstallprompt",e=>{
  e.preventDefault();
  ssrInstallPrompt=e;
  if(installBtn) installBtn.hidden=false;
  if(installGuide) installGuide.textContent="Android：下の「アプリをインストール」を押してください。";
});
if(installBtn) installBtn.addEventListener("click",async()=>{
  if(!ssrInstallPrompt)return;
  await ssrInstallPrompt.prompt();
  ssrInstallPrompt=null;
  installBtn.hidden=true;
});
window.addEventListener("appinstalled",()=>{
  if(installBtn) installBtn.hidden=true;
  if(installGuide) installGuide.textContent="インストール済みです。ホーム画面の「SSR温測」アイコンから起動できます。";
});
const isiOS=/iphone|ipad|ipod/i.test(navigator.userAgent);
const standalone=(window.matchMedia&&window.matchMedia("(display-mode: standalone)").matches)||navigator.standalone===true;
if(isiOS&&!standalone&&installGuide){
  installGuide.textContent="iPhone・iPad：Safariで開き、共有ボタン →「ホーム画面に追加」→ 追加。専用アイコンから起動できます。";
}

const androidGuide=document.getElementById("androidGuide");
const iosGuide=document.getElementById("iosGuide");
const isAndroid=/android/i.test(navigator.userAgent);
if(isiOS && iosGuide) iosGuide.classList.add("current");
if(isAndroid && androidGuide) androidGuide.classList.add("current");
if(standalone){
  document.body.classList.add("app-mode");
  const installBox=document.getElementById("installBox");
  if(installBox) installBox.hidden=true;
  if(installBtn) installBtn.hidden=true;
}

const missingBox=document.getElementById("missingBox");
const missingList=document.getElementById("missingList");

function machineBlockRange(machine){
  const lines=input.value.replace(/\r/g,"").split("\n");
  let start=-1;
  let end=lines.length;

  for(let i=0;i<lines.length;i++){
    const hit=lines[i].match(MACHINE_RE);
    if(hit&&normalizeMachineName(hit[1])===machine){
      start=i;
      break;
    }
  }
  if(start<0)return null;

  for(let i=start+1;i<lines.length;i++){
    if(lines[i].match(MACHINE_RE)){
      end=i;
      break;
    }
  }
  return {lines,start,end};
}

function expectedRowSpecs(machine){
  return [
    {label:"盤内温度",count:2},
    ...(MACHINE_LAYOUT[machine]||[]).map((count,i)=>({
      label:"SSR "+(i+1)+"段目",
      count
    }))
  ];
}

function machineDataRows(machine){
  const range=machineBlockRange(machine);
  if(!range)return null;

  const rows=[];
  const machineLineText=range.lines[range.start]||"";
  const trailing=machineLineText.replace(MACHINE_RE," ").trim();
  if(trailing && /\d/.test(trailing)){
    rows.push({
      lineIndex:range.start,
      text:trailing,
      onMachineLine:true
    });
  }

  for(let i=range.start+1;i<range.end;i++){
    const text=range.lines[i].trim();
    if(!text || !/\d/.test(text))continue;
    rows.push({
      lineIndex:i,
      text:range.lines[i],
      onMachineLine:false
    });
  }

  return {range,rows};
}

function analyzeMachineRows(machine){
  const src=machineDataRows(machine);
  if(!src)return [];

  const specs=expectedRowSpecs(machine);
  const issues=[];

  for(let i=0;i<specs.length;i++){
    const spec=specs[i];
    const row=src.rows[i]||null;
    const text=row?row.text:"";
    const parsed=parseMeasurementValues(text);

    const validValues=parsed.values.filter(v=>/^-?\d{1,3}\.\d$/.test(v));
    const bad=parsed.invalid || validValues.length!==spec.count;

    if(bad){
      issues.push({
        machine,
        specIndex:i,
        label:spec.label,
        expected:spec.count,
        row,
        text,
        parsedCount:validValues.length
      });
    }
  }

  return issues;
}

function replaceMachineRow(machine,specIndex,newText){
  const src=machineDataRows(machine);
  if(!src)return false;

  const row=src.rows[specIndex];
  const clean=newText.replace(/\r/g,"").trim();

  if(row){
    if(row.onMachineLine){
      const original=src.range.lines[src.range.start];
      const match=original.match(MACHINE_RE);
      const name=match?match[0]:machine;
      src.range.lines[src.range.start]=name+(clean?" "+clean:"");
    }else{
      src.range.lines[row.lineIndex]=clean;
    }
  }else{
    let insertAt=src.range.end;
    if(src.rows.length){
      insertAt=Math.min(src.range.end,src.rows[src.rows.length-1].lineIndex+1);
    }else{
      insertAt=src.range.start+1;
    }
    src.range.lines.splice(insertAt,0,clean);
  }

  input.value=src.range.lines.join("\n");
  return true;
}

function renderMissingData(){
  if(!missingBox||!missingList)return 0;
  missingList.replaceChildren();

  let total=0;
  const seen=new Set();

  for(const block of parseMachineBlocks(input.value)){
    const machine=block.machine;
    if(seen.has(machine))continue;
    seen.add(machine);

    const issues=analyzeMachineRows(machine);
    if(!issues.length)continue;

    for(const issue of issues){
      total++;

      const card=document.createElement("div");
      card.className="stable-correction-card";

      const title=document.createElement("div");
      title.className="stable-correction-title";
      title.textContent=machine+"　"+issue.label+"　要修正";
      card.appendChild(title);

      const note=document.createElement("div");
      note.className="stable-correction-note";
      note.textContent="このWord行だけを修正してください。";
      card.appendChild(note);

      const edit=document.createElement("input");
      edit.type="text";
      edit.className="stable-correction-input line-correction-input";
      edit.value=issue.text;
      edit.placeholder="この行を修正";
      edit.autocomplete="off";
      edit.spellcheck=false;
      card.appendChild(edit);

      const apply=document.createElement("button");
      apply.type="button";
      apply.className="pressable stable-correction-apply";
      apply.textContent="この行を反映";
      apply.addEventListener("click",()=>{
        if(replaceMachineRow(machine,issue.specIndex,edit.value)){
          pendingGenerate=false;
          const remaining=renderMissingData();
          result.textContent=remaining
            ? machine+" "+issue.label+" を反映しました。残りの要修正行を確認してください。"
            : "修正完了。Excelを保存できます。";
        }
      });

      card.appendChild(apply);
      missingList.appendChild(card);
    }
  }

  missingBox.hidden=total===0;
  return total;
}

input.addEventListener("input",()=>{
  pendingGenerate=false;
  renderMissingData();
});
input.addEventListener("change",()=>{
  pendingGenerate=false;
  renderMissingData();
});
renderMissingData();

