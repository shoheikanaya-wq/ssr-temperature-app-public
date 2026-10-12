const distributionToggle=document.getElementById("distributionToggle");
const distributionPanel=document.getElementById("distributionPanel");
const distributionClose=document.getElementById("distributionClose");

function setDistributionOpen(open){
  if(!distributionPanel||!distributionToggle)return;
  distributionPanel.hidden=!open;
  distributionToggle.setAttribute("aria-expanded",open?"true":"false");
  if(open)distributionPanel.scrollIntoView({behavior:"smooth",block:"nearest"});
}

distributionToggle?.addEventListener("click",()=>{
  setDistributionOpen(distributionPanel?.hidden!==false);
});
distributionClose?.addEventListener("click",()=>setDistributionOpen(false));
