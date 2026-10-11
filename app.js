const VERSION="3.25";
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
      refreshInputSuggestion();
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
      const after=raw.replace(hit[0]," ");
      current={machine,text:after,count:countNumbers(after)};
    }else if(current){
      current.text+="\n"+raw;
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

function findValidationIssue(text){
  return findInputSuggestion(text)||findCountIssue(text);
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
  const a=document.createElement("a");a.href=URL.createObjectURL(out);a.download="SSR温測_v3.25_生成.xlsx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
  result.textContent="生成完了。\n入力 "+lines.length+" 行を「スマホ入力」へ反映しました。\nExcelを開いて横・縦データを確認してください。";
 }catch(e){result.textContent="エラー: "+e.message}
}
let pendingGenerate=false;
document.getElementById("generate").addEventListener("click",()=>{
  const joined=findInputSuggestion(input.value);
  if(joined){
    pendingGenerate=true;
    refreshInputSuggestion();
    result.textContent="怪しい入力を確認してください。";
    assistBox?.scrollIntoView({behavior:"smooth",block:"center"});
    return;
  }
  const countIssue=findCountIssue(input.value);
  if(countIssue){
    pendingGenerate=true;
    renderMissingData();
    result.textContent="不足データを手入力してください。";
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

const assistBox=document.getElementById("assistBox");
const assistText=document.getElementById("assistText");
const assistChoices=document.getElementById("assistChoices");
const assistManual=document.getElementById("assistManual");

const missingBox=document.getElementById("missingBox");
const missingList=document.getElementById("missingList");

function machineSlots(machine){
  const rows=MACHINE_LAYOUT[machine]||[];
  const slots=[
    {index:0,label:"盤内温度 上"},
    {index:1,label:"盤内温度 下"}
  ];
  let idx=2;
  rows.forEach((count,row)=>{
    for(let pos=1;pos<=count;pos++){
      slots.push({index:idx,label:"SSR "+(row+1)+"段目 "+pos+"番"});
      idx++;
    }
  });
  return slots;
}

function machineRange(text,machine){
  const re=new RegExp(MACHINE_RE.source,"ig");
  let m;
  while((m=re.exec(text))){
    if(m[1].toUpperCase()!==machine)continue;
    const start=m.index+m[0].length;
    const next=re.exec(text);
    return {start,end:next?next.index:text.length};
  }
  return null;
}

function insertMissingValue(machine,slotIndex,value){
  const range=machineRange(input.value,machine);
  if(!range)return false;
  const part=input.value.slice(range.start,range.end);
  const nums=[...part.matchAll(/-?\d+(?:\.\d+)?/g)];
  let at=range.end;
  let prefix="\n";
  if(slotIndex<nums.length){
    at=range.start+nums[slotIndex].index;
    prefix="";
  }
  input.value=input.value.slice(0,at)+prefix+value+" "+input.value.slice(at);
  return true;
}

function renderMissingData(){
  if(!missingBox||!missingList)return;
  missingList.replaceChildren();
  const issues=parseMachineBlocks(input.value)
    .map(block=>({...block,expected:EXPECTED_COUNTS[block.machine]||0}))
    .filter(block=>block.expected&&block.count<block.expected);

  if(!issues.length){
    missingBox.hidden=true;
    return;
  }

  for(const issue of issues){
    const card=document.createElement("div");
    card.className="missing-card";

    const title=document.createElement("div");
    title.className="missing-title";
    title.textContent=issue.machine+"：あと"+(issue.expected-issue.count)+"個不足";
    card.appendChild(title);

    const rows=MACHINE_LAYOUT[issue.machine];
    const detail=document.createElement("div");
    detail.className="missing-detail";
    detail.textContent="盤内温度2個 / SSR "+rows.map((n,i)=>(i+1)+"段目 "+n+"個").join(" / ");
    card.appendChild(detail);

    const entry=document.createElement("div");
    entry.className="missing-entry";

    const select=document.createElement("select");
    select.className="missing-position";
    for(const slot of machineSlots(issue.machine)){
      const o=document.createElement("option");
      o.value=String(slot.index);
      o.textContent=slot.label;
      if(slot.index===Math.min(issue.count,issue.expected-1))o.selected=true;
      select.appendChild(o);
    }

    const field=document.createElement("input");
    field.type="number";
    field.step="0.1";
    field.inputMode="decimal";
    field.placeholder="不足値";

    const add=document.createElement("button");
    add.type="button";
    add.className="pressable missing-add";
    add.textContent="追加";
    add.addEventListener("click",()=>{
      const value=field.value.trim();
      if(!value){field.focus();return;}
      if(!insertMissingValue(issue.machine,Number(select.value),value))return;
      refreshInputSuggestion();
      renderMissingData();
      const still=findCountIssue(input.value);
      if(pendingGenerate&&!still&&!findInputSuggestion(input.value)){
        pendingGenerate=false;
        result.textContent="確認完了。Excelを生成します…";
        build();
      }else{
        result.textContent=issue.machine+" に "+value+" を追加しました。";
      }
    });

    entry.append(select,field,add);
    card.appendChild(entry);
    missingList.appendChild(card);
  }
  missingBox.hidden=false;
}

function findInputSuggestion(text){
  const chunks=text.split(/\s+/).filter(Boolean);
  for(const chunk of chunks){
    let m=chunk.match(/^(\d{2}\.\d)(\d{2}\.\d)(\d\.\d)$/);
    if(m){
      return {
        chunk,
        message:"「"+chunk+"」は数値がつながっている可能性があります。どれですか？",
        choices:[
          m[1]+" "+m[2]+" 3"+m[3],
          m[1]+" "+m[2]+" 4"+m[3],
          m[1]+" "+m[2]+" 5"+m[3]
        ]
      };
    }
    m=chunk.match(/^(\d{2}\.\d)(\d{2}\.\d)$/);
    if(m){
      return {
        chunk,
        message:"「"+chunk+"」は2つの測定値がつながっている可能性があります。",
        choices:[m[1]+" "+m[2]]
      };
    }
  }
  return null;
}

function refreshInputSuggestion(){
  if(!assistBox||!assistText||!assistChoices)return;
  const s=findInputSuggestion(input.value);
  assistChoices.replaceChildren();
  if(!s){
    assistBox.hidden=true;
    return;
  }
  assistText.textContent=s.message;
  for(const choice of s.choices||[]){
    const b=document.createElement("button");
    b.type="button";
    b.className="pressable assist-choice";
    b.textContent=choice;
    b.addEventListener("click",()=>{
      input.value=input.value.replace(s.chunk,choice);
      input.focus();
      const next=findInputSuggestion(input.value);
      refreshInputSuggestion();
      if(pendingGenerate && !next){
        pendingGenerate=false;
        result.textContent="確認完了。Excelを生成します…";
        build();
      }else if(pendingGenerate && next){
        result.textContent="次の怪しい入力を確認してください。";
        assistBox?.scrollIntoView({behavior:"smooth",block:"center"});
      }
    });
    assistChoices.appendChild(b);
  }
  if(s.type==="count" && (!s.choices || !s.choices.length)){
    assistManual.textContent="手入力で確認";
  }else if(assistManual){
    assistManual.textContent="どれでもない";
  }
  assistBox.hidden=false;
}

input.addEventListener("input",()=>{
  pendingGenerate=false;
  refreshInputSuggestion();
  renderMissingData();
});
input.addEventListener("change",()=>{
  pendingGenerate=false;
  refreshInputSuggestion();
  renderMissingData();
});
if(assistManual)assistManual.addEventListener("click",()=>{
  pendingGenerate=false;
  assistBox.hidden=true;
  input.focus();
  result.textContent="候補を使わず手入力で修正してください。";
});
refreshInputSuggestion();
renderMissingData();
