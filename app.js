const VERSION="3.33";
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
  const a=document.createElement("a");a.href=URL.createObjectURL(out);a.download="SSR温測_v3.33_生成.xlsx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
  result.textContent="生成完了。\n入力 "+lines.length+" 行を「スマホ入力」へ反映しました。\nExcelを開いて横・縦データを確認してください。";
 }catch(e){result.textContent="エラー: "+e.message}
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

function expectedRows(machine){
  return [
    {label:"盤内温度",count:2,type:"panel"},
    ...(MACHINE_LAYOUT[machine]||[]).map((count,i)=>({label:"SSR "+(i+1)+"段目",count,type:"ssr"}))
  ];
}

function sourceBlock(machine){
  const lines=input.value.replace(/\r/g,"").split("\n");
  let machineLine=-1,end=lines.length;
  for(let i=0;i<lines.length;i++){
    const hit=lines[i].match(MACHINE_RE);
    if(hit&&hit[1].toUpperCase()===machine){machineLine=i;break;}
  }
  if(machineLine<0)return null;
  for(let i=machineLine+1;i<lines.length;i++){
    if(lines[i].match(MACHINE_RE)){end=i;break;}
  }
  const rows=[];
  const sameLine=lines[machineLine].replace(MACHINE_RE,"").trim();
  if(sameLine)rows.push({lineIndex:machineLine,text:sameLine,onMachineLine:true});
  for(let i=machineLine+1;i<end;i++){
    if(lines[i].trim())rows.push({lineIndex:i,text:lines[i],onMachineLine:false});
  }
  return {lines,machineLine,end,rows};
}

function splitRowValues(line){
  return parseMeasurementValues(line).values;
}

function parsePanelTemperatures(line){
  const parsed=parseMeasurementValues(line);
  const values=parsed.values;
  const valid=!parsed.invalid &&
    values.length===2 &&
    values.every(v=>/^\d{2}\.\d$/.test(v));
  return {valid,values};
}

function analyzeRow(line,spec){
  if(spec.type==="panel"){
    const panel=parsePanelTemperatures(line);
    if(panel.valid)return {parts:splitRowValues(line),issues:[],logicalCount:2};

    const parts=splitRowValues(line);
    return {
      parts,
      logicalCount:panel.values.length,
      issues:[{
        kind:"replace-line",
        index:0,
        label:"要修正",
        current:line.trim(),
        placeholder:"例：32.0 28.9"
      }]
    };
  }

  const expectedCount=spec.count;
  const parsed=parseMeasurementValues(line);
  const parts=parsed.values;
  const malformed=[];
  parts.forEach((p,i)=>{
    if(!/^-?\d{1,3}(?:\.\d)?$/.test(p))malformed.push(i);
  });
  if(parsed.invalid && !malformed.length && parts.length){
    malformed.push(parts.length-1);
  }

  const issues=[];
  for(const i of malformed){
    issues.push({
      kind:"replace",
      index:i,
      label:"ここがおかしい？",
      current:parts[i]||"",
      placeholder:"この部分だけ修正"
    });
  }

  return {parts,issues,logicalCount:parts.length};
}

function renderHighlightedLine(line,spec){
  const {parts,issues}=analyzeRow(line,spec);
  const expectedCount=spec.count;
  const badIndexes=new Map();
  issues.forEach(issue=>badIndexes.set(issue.index,issue.label));

  const preview=document.createElement("div");
  preview.className="raw-line-preview";

  if(!parts.length){
    const s=document.createElement("span");
    s.className="bad-part";
    s.textContent="（この行が空です）";
    preview.appendChild(s);
    return preview;
  }

  parts.forEach((p,i)=>{
    if(i)preview.append(" ");
    const span=document.createElement("span");
    span.textContent=p;
    if(badIndexes.has(i)){
      span.className="bad-part";
      span.title=badIndexes.get(i);
    }
    preview.appendChild(span);
  });

  return preview;
}

function replaceIssueInSource(machine,rowIndex,issueIndex,newText,replaceWholeLine=false){
  const src=sourceBlock(machine);
  if(!src)return false;
  const row=src.rows[rowIndex];
  if(!row)return false;

  const parts=splitRowValues(row.text);
  const replacement=newText.trim().split(/\s+/).filter(Boolean);

  let rebuilt;
  if(replaceWholeLine){
    rebuilt=newText.trim();
  }else{
    if(issueIndex<parts.length){
      parts.splice(issueIndex,1,...replacement);
    }else{
      parts.push(...replacement);
    }
    rebuilt=parts.join(" ");
  }
  if(row.onMachineLine){
    const name=(src.lines[src.machineLine].match(MACHINE_RE)||[machine])[0];
    src.lines[src.machineLine]=name+(rebuilt?" "+rebuilt:"");
  }else{
    src.lines[row.lineIndex]=rebuilt;
  }
  input.value=src.lines.join("\n");
  return true;
}

function renderMissingData(){
  if(!missingBox||!missingList)return;
  missingList.replaceChildren();
  let total=0;

  for(const block of parseMachineBlocks(input.value)){
    const machine=block.machine;
    if(!MACHINE_LAYOUT[machine])continue;
    const src=sourceBlock(machine);
    if(!src)continue;
    const specs=expectedRows(machine);

    const rowsToShow=[];
    for(let i=0;i<specs.length;i++){
      const row=src.rows[i]||{text:"",lineIndex:null,onMachineLine:false};
      const analysis=analyzeRow(row.text,specs[i]);
      if(analysis.issues.length){
        rowsToShow.push({index:i,row,spec:specs[i],analysis});
      }
    }
    if(!rowsToShow.length)continue;

    const card=document.createElement("div");
    card.className="missing-card";

    const title=document.createElement("div");
    title.className="missing-title";
    title.textContent=machine+"　修正箇所";
    card.appendChild(title);

    for(const item of rowsToShow){
      total++;
      const wrap=document.createElement("div");
      wrap.className="raw-row-editor";

      const head=document.createElement("div");
      head.className="raw-row-head";
      head.textContent=item.spec.label+"　要修正";
      wrap.appendChild(head);

      wrap.appendChild(renderHighlightedLine(item.row.text,item.spec));

      for(const issue of item.analysis.issues){
        const correction=document.createElement("div");
        correction.className="issue-editor";

        const label=document.createElement("div");
        label.className="issue-label";
        label.textContent=issue.label;
        correction.appendChild(label);

        const edit=document.createElement("input");
        edit.type="text";
        edit.className="issue-input";
        edit.value=issue.current;
        edit.placeholder=issue.placeholder;
        edit.autocomplete="off";
        edit.spellcheck=false;
        correction.appendChild(edit);

        const apply=document.createElement("button");
        apply.type="button";
        apply.className="pressable issue-apply";
        apply.textContent="修正を反映";
        apply.addEventListener("click",()=>{
          if(replaceIssueInSource(machine,item.index,issue.index,edit.value,issue.kind==="replace-line")){
            pendingGenerate=false;
            renderMissingData();
            result.textContent=machine+" "+item.spec.label+" の要修正箇所を反映しました。";
          }
        });
        correction.appendChild(apply);
        wrap.appendChild(correction);
      }

      card.appendChild(wrap);
    }

    missingList.appendChild(card);
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

