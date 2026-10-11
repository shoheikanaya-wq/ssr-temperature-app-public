const VERSION="3.27";
const input=document.getElementById("input"),template=document.getElementById("template"),result=document.getElementById("result");
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
const MACHINE_RE=/(FM83-1|FM83-2|FM81|FM84|FM96|FM97|FM19|FM22|FM01|FM07|FE01|FE04|MM18)/i;

function countNumbers(text){
  const m=text.match(/-?\d+(?:\.\d+)?/g);
  return m?m.length:0;
}

function parseMachineBlocks(text){
  const blocks=[];
  let current=null;
  for(const raw of text.replace(/\r/g,"").split("\n")){
    const hit=raw.match(MACHINE_RE);
    if(hit){
      if(current)blocks.push(current);
      const machine=hit[1].toUpperCase();
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

function findCountIssue(text){
  for(const block of parseMachineBlocks(text)){
    const expected=EXPECTED_COUNTS[block.machine];
    if(!expected)continue;
    if(block.count!==expected){
      const diff=expected-block.count;
      return {
        type:"count",
        machine:block.machine,
        expected,
        actual:block.count,
        chunk:null,
        choices:[],
        message:diff>0
          ? block.machine+"：必要 "+expected+"値に対して "+block.count+"値です。あと"+diff+"個不足しています。"
          : block.machine+"：必要 "+expected+"値に対して "+block.count+"値です。"+Math.abs(diff)+"個多いので確認してください。"
      };
    }
  }
  return null;
}

async function build(){
 try{
  const text=input.value.replace(/\r/g,"").trim();
  if(!text)throw new Error("Word測定データを貼り付けてください。");
  if(!template.files.length)throw new Error("基準Excel v3.22を選択してください。");
  result.textContent="生成中…";
  const zip=await JSZip.loadAsync(await template.files[0].arrayBuffer());
  const parser=new DOMParser(),ser=new XMLSerializer();
  const wbxml=await zip.file("xl/workbook.xml").async("string");
  const wbdoc=parser.parseFromString(wbxml,"application/xml");
  const sheet=[...wbdoc.getElementsByTagName("sheet")].find(x=>x.getAttribute("name")==="スマホ入力");
  if(!sheet)throw new Error("「スマホ入力」シートが見つかりません。基準Excelを確認してください。");
  const rid=sheet.getAttribute("r:id");
  const relxml=await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const reldoc=parser.parseFromString(relxml,"application/xml");
  const rel=[...reldoc.getElementsByTagName("Relationship")].find(x=>x.getAttribute("Id")===rid);
  if(!rel)throw new Error("スマホ入力シートの参照を取得できません。");
  let target=rel.getAttribute("Target").replace(/^\//,"");
  if(!target.startsWith("xl/"))target="xl/"+target.replace(/^\.\//,"");
  const sxml=await zip.file(target).async("string");
  const sdoc=parser.parseFromString(sxml,"application/xml");
  const ns=sdoc.documentElement.namespaceURI;
  const sheetData=sdoc.getElementsByTagName("sheetData")[0];
  const rows=[...sheetData.getElementsByTagName("row")];
  const rowMap=new Map(rows.map(r=>[Number(r.getAttribute("r")),r]));
  for(let n=5;n<=204;n++){
   let row=rowMap.get(n);
   if(!row){row=sdoc.createElementNS(ns,"row");row.setAttribute("r",String(n));sheetData.appendChild(row);rowMap.set(n,row)}
   for(const c of [...row.getElementsByTagName("c")])if(c.getAttribute("r")===colA(n))row.removeChild(c);
  }
  const lines=text.split("\n").slice(0,200);
  lines.forEach((line,i)=>{
   const n=5+i,row=rowMap.get(n);
   const c=sdoc.createElementNS(ns,"c");c.setAttribute("r",colA(n));c.setAttribute("t","inlineStr");
   const is=sdoc.createElementNS(ns,"is"),t=sdoc.createElementNS(ns,"t");t.setAttribute("xml:space","preserve");t.textContent=line;
   is.appendChild(t);c.appendChild(is);row.appendChild(c);
  });
  zip.file(target,ser.serializeToString(sdoc));
  let calc=wbdoc.getElementsByTagName("calcPr")[0];
  if(!calc){calc=wbdoc.createElementNS(wbdoc.documentElement.namespaceURI,"calcPr");wbdoc.documentElement.appendChild(calc)}
  calc.setAttribute("fullCalcOnLoad","1");calc.setAttribute("forceFullCalc","1");calc.setAttribute("calcMode","auto");
  zip.file("xl/workbook.xml",ser.serializeToString(wbdoc));
  const out=await zip.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement("a");a.href=URL.createObjectURL(out);a.download="SSR温測_v3.27_生成.xlsx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
  result.textContent="生成完了。\n入力 "+lines.length+" 行を「スマホ入力」へ反映しました。\nExcelを開いて横・縦データを確認してください。";
 }catch(e){result.textContent="エラー: "+e.message}
}
let pendingGenerate=false;
document.getElementById("generate").addEventListener("click",()=>{
  const countIssue=findCountIssue(input.value);
  if(countIssue){
    pendingGenerate=true;
    renderMissingData();
    result.textContent="不足している行を手入力してください。";
    missingBox?.scrollIntoView({behavior:"smooth",block:"center"});
    return;
  }
  pendingGenerate=false;
  build();
});
if("serviceWorker" in navigator){
  let refreshing=false;
  const updateBox=document.getElementById("updateBox");
  const updateBtn=document.getElementById("updateApp");
  const showUpdate=reg=>{
    if(!reg.waiting)return;
    if(updateBox)updateBox.hidden=false;
    if(updateBtn)updateBtn.onclick=()=>{
      updateBtn.disabled=true;
      reg.waiting.postMessage({type:"SKIP_WAITING"});
    };
  };
  navigator.serviceWorker.addEventListener("controllerchange",()=>{
    if(refreshing)return;
    refreshing=true;
    location.reload();
  });
  navigator.serviceWorker.register("./sw.js").then(reg=>{
    showUpdate(reg);
    reg.addEventListener("updatefound",()=>{
      const worker=reg.installing;
      if(!worker)return;
      worker.addEventListener("statechange",()=>{
        if(worker.state==="installed" && navigator.serviceWorker.controller)showUpdate(reg);
      });
    });
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

function rowSpec(machine){
  return [
    {label:"盤内温度",count:2},
    ...(MACHINE_LAYOUT[machine]||[]).map((count,i)=>({
      label:"SSR "+(i+1)+"段目",
      count
    }))
  ];
}

function numericChunks(line){
  return line
    .trim()
    .split(/\s+/)
    .map(s=>s.replace(/^[,、;；]+|[,、;；]+$/g,""))
    .filter(s=>/\d/.test(s));
}

function validMeasurement(value){
  return /^-?\d{1,3}(?:\.\d)?$/.test(value);
}

function machineSource(machine){
  const lines=input.value.replace(/\r/g,"").split("\n");
  let start=-1,end=lines.length;
  for(let i=0;i<lines.length;i++){
    const hit=lines[i].match(MACHINE_RE);
    if(hit&&hit[1].toUpperCase()===machine){
      start=i;
      break;
    }
  }
  if(start<0)return null;
  for(let i=start+1;i<lines.length;i++){
    if(MACHINE_RE.test(lines[i])){
      end=i;
      break;
    }
  }
  const dataRows=[];
  const first=lines[start].replace(MACHINE_RE," ").trim();
  if(/\d/.test(first))dataRows.push({lineIndex:start,text:first,onMachineLine:true});
  for(let i=start+1;i<end;i++){
    if(/\d/.test(lines[i]))dataRows.push({lineIndex:i,text:lines[i],onMachineLine:false});
  }
  return {lines,start,end,dataRows};
}

function updateSourceRow(machine,specIndex,values){
  const src=machineSource(machine);
  if(!src)return;
  const specs=rowSpec(machine);
  const normalized=values.map(v=>v.trim()).filter(v=>v!=="").join(" ");
  const row=src.dataRows[specIndex];

  if(row){
    if(row.onMachineLine){
      const nameMatch=src.lines[src.start].match(MACHINE_RE);
      const name=nameMatch?nameMatch[0]:machine;
      src.lines[src.start]=name+(normalized?" "+normalized:"");
    }else{
      src.lines[row.lineIndex]=normalized;
    }
  }else{
    let insertAt=src.end;
    if(src.dataRows.length){
      insertAt=Math.min(src.end,src.dataRows[src.dataRows.length-1].lineIndex+1);
    }else{
      insertAt=src.start+1;
    }
    src.lines.splice(insertAt,0,normalized);
  }
  input.value=src.lines.join("\n");
}

function renderEditableRow(machine,specIndex,spec,row){
  const chunks=row?numericChunks(row.text):[];
  const malformed=chunks.map(v=>!validMeasurement(v));
  const wrongCount=chunks.length!==spec.count;
  const hasMalformed=malformed.some(Boolean);
  if(!wrongCount&&!hasMalformed)return null;

  const wrap=document.createElement("div");
  wrap.className="missing-row full-row-edit";
  if(wrongCount)wrap.classList.add("row-problem");

  const head=document.createElement("div");
  head.className="missing-row-head";
  head.textContent=spec.label+"　"+chunks.length+"/"+spec.count;
  wrap.appendChild(head);

  const note=document.createElement("div");
  note.className="row-problem-note";
  if(hasMalformed){
    note.textContent="赤い値が怪しいです";
  }else if(chunks.length<spec.count){
    note.textContent="この行で "+(spec.count-chunks.length)+"個不足";
  }else{
    note.textContent="この行で "+(chunks.length-spec.count)+"個多い";
  }
  wrap.appendChild(note);

  const grid=document.createElement("div");
  grid.className="missing-grid";
  const fieldCount=Math.max(spec.count,chunks.length);

  for(let i=0;i<fieldCount;i++){
    const cell=document.createElement("div");
    cell.className="missing-cell";

    const label=document.createElement("span");
    label.className="missing-cell-label";
    label.textContent=specIndex===0?(i===0?"上":i===1?"下":String(i+1)):String(i+1);

    const field=document.createElement("input");
    field.type="text";
    field.inputMode="decimal";
    field.value=chunks[i]??"";
    field.placeholder="—";
    field.dataset.pos=String(i);

    const suspicious=malformed[i]===true || (i>=chunks.length && wrongCount);
    if(suspicious){
      cell.classList.add("likely-missing");
      const tag=document.createElement("span");
      tag.className="likely-tag";
      tag.textContent=malformed[i]===true?"ここがおかしい？":"不足";
      cell.appendChild(tag);
    }

    cell.append(label,field);
    grid.appendChild(cell);
  }
  wrap.appendChild(grid);

  const apply=document.createElement("button");
  apply.type="button";
  apply.className="pressable row-apply";
  apply.textContent="この行を反映";
  apply.addEventListener("click",()=>{
    const values=[...grid.querySelectorAll("input")].map(x=>x.value);
    updateSourceRow(machine,specIndex,values);
    pendingGenerate=false;
    renderMissingData();
    result.textContent=machine+" "+spec.label+" を反映しました。";
  });
  wrap.appendChild(apply);

  return wrap;
}

function renderMissingData(){
  if(!missingBox||!missingList)return;
  missingList.replaceChildren();
  let problemCount=0;

  for(const block of parseMachineBlocks(input.value)){
    const machine=block.machine;
    if(!MACHINE_LAYOUT[machine])continue;
    const src=machineSource(machine);
    if(!src)continue;

    const specs=rowSpec(machine);
    const card=document.createElement("div");
    card.className="missing-card";
    card.dataset.machine=machine;

    const rows=[];
    for(let i=0;i<specs.length;i++){
      const rendered=renderEditableRow(machine,i,specs[i],src.dataRows[i]);
      if(rendered){
        rows.push(rendered);
        problemCount++;
      }
    }
    if(!rows.length)continue;

    const title=document.createElement("div");
    title.className="missing-title";
    title.textContent=machine+"　修正が必要な行";
    card.appendChild(title);
    rows.forEach(row=>card.appendChild(row));
    missingList.appendChild(card);
  }

  missingBox.hidden=problemCount===0;
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

