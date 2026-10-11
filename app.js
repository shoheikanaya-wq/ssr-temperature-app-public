const VERSION="1.10";
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
  result.textContent="生成中…";
  const templateResponse=await fetch("./SSR温測テンプレ.xlsx",{cache:"no-store"});
  if(!templateResponse.ok)throw new Error("標準ひな形を読み込めませんでした。");
  const zip=await JSZip.loadAsync(await templateResponse.arrayBuffer());
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
  const a=document.createElement("a");a.href=URL.createObjectURL(out);a.download="SSR温測_生成.xlsx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
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

function machineBlockRange(machine){
  const lines=input.value.replace(/\r/g,"").split("\n");
  let start=-1;
  let end=lines.length;

  for(let i=0;i<lines.length;i++){
    const hit=lines[i].match(MACHINE_RE);
    if(hit&&hit[1].toUpperCase()===machine){
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

function analyzeMachineBlock(machine){
  const range=machineBlockRange(machine);
  if(!range)return null;

  const rawLines=range.lines.slice(range.start,range.end);
  const dataLines=[...rawLines];
  if(dataLines.length){
    dataLines[0]=dataLines[0].replace(MACHINE_RE," ").trim();
  }

  let count=0;
  let invalid=false;
  for(const line of dataLines){
    const parsed=parseMeasurementValues(line);
    count+=parsed.values.length;
    if(parsed.invalid)invalid=true;
  }

  const expected=EXPECTED_COUNTS[machine]||0;
  return {
    machine,
    range,
    raw:rawLines.join("\n"),
    count,
    expected,
    invalid,
    issue:invalid || (count>0 && expected>0 && count!==expected)
  };
}

function replaceMachineBlock(machine,newText){
  const range=machineBlockRange(machine);
  if(!range)return false;

  const replacement=newText.replace(/\r/g,"").split("\n");
  range.lines.splice(range.start,range.end-range.start,...replacement);
  input.value=range.lines.join("\n");
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

    const analysis=analyzeMachineBlock(machine);
    if(!analysis || !analysis.issue)continue;
    total++;

    const card=document.createElement("div");
    card.className="stable-correction-card";

    const title=document.createElement("div");
    title.className="stable-correction-title";
    title.textContent=machine+"　要修正";
    card.appendChild(title);

    const note=document.createElement("div");
    note.className="stable-correction-note";
    note.textContent="Word読取元データをそのまま修正してください。";
    card.appendChild(note);

    const edit=document.createElement("textarea");
    edit.className="stable-correction-input";
    edit.rows=Math.max(4,analysis.raw.split("\n").length+1);
    edit.value=analysis.raw;
    edit.spellcheck=false;
    card.appendChild(edit);

    const apply=document.createElement("button");
    apply.type="button";
    apply.className="pressable stable-correction-apply";
    apply.textContent="修正を反映";
    apply.addEventListener("click",()=>{
      if(replaceMachineBlock(machine,edit.value)){
        pendingGenerate=false;
        const remaining=renderMissingData();
        result.textContent=remaining
          ? machine+" を反映しました。残りの要修正データを確認してください。"
          : "修正完了。Excelを保存できます。";
      }
    });
    card.appendChild(apply);
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

