const VERSION="3.23";
const input=document.getElementById("input"),template=document.getElementById("template"),result=document.getElementById("result");
const esc=s=>s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
function colA(row){return "A"+row}
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
  const a=document.createElement("a");a.href=URL.createObjectURL(out);a.download="SSR温測_v3.23_生成.xlsx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),3000);
  result.textContent="生成完了。\n入力 "+lines.length+" 行を「スマホ入力」へ反映しました。\nExcelを開いて横・縦データを確認してください。";
 }catch(e){result.textContent="エラー: "+e.message}
}
let pendingGenerate=false;
document.getElementById("generate").addEventListener("click",()=>{
  const suspicious=findInputSuggestion(input.value);
  if(suspicious){
    pendingGenerate=true;
    refreshInputSuggestion();
    result.textContent="怪しい入力を確認してください。候補をタップすると続けます。";
    assistBox?.scrollIntoView({behavior:"smooth",block:"center"});
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
if(standalone && installGuide){
  installGuide.textContent="この端末にはSSR温測が追加済みです。";
  if(installBtn) installBtn.hidden=true;
}

const assistBox=document.getElementById("assistBox");
const assistText=document.getElementById("assistText");
const assistChoices=document.getElementById("assistChoices");
const assistManual=document.getElementById("assistManual");

function findInputSuggestion(text){
  const tokens=text.split(/\s+/).filter(Boolean);
  for(const token of tokens){
    let m=token.match(/^(\d{2}\.\d)(\d{2}\.\d)(\d\.\d)$/);
    if(m){
      return {
        token,
        message:"「"+token+"」は数値がつながっている可能性があります。どれですか？",
        choices:[
          m[1]+" "+m[2]+" 3"+m[3],
          m[1]+" "+m[2]+" 4"+m[3],
          m[1]+" "+m[2]+" 5"+m[3]
        ]
      };
    }
    m=token.match(/^(\d{2}\.\d)(\d{2}\.\d)$/);
    if(m){
      return {
        token,
        message:"「"+token+"」は2つの測定値がつながっている可能性があります。",
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
  for(const choice of s.choices){
    const b=document.createElement("button");
    b.type="button";
    b.className="pressable assist-choice";
    b.textContent=choice;
    b.addEventListener("click",()=>{
      input.value=input.value.replace(s.token,choice);
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
  assistBox.hidden=false;
}

input.addEventListener("input",()=>{
  pendingGenerate=false;
  refreshInputSuggestion();
});
input.addEventListener("change",()=>{
  pendingGenerate=false;
  refreshInputSuggestion();
});
if(assistManual)assistManual.addEventListener("click",()=>{
  pendingGenerate=false;
  assistBox.hidden=true;
  input.focus();
  result.textContent="候補を使わず手入力で修正してください。";
});
refreshInputSuggestion();
