const $=id=>document.getElementById(id);
let rawRows=[], mappings=[];

// Client-verified mapping from Book2.xlsx -> Sheet1.
// IMPORTANT: this mapping is used for campaign attribution only.
// Answered/drop classification rules remain separate below.
const DEFAULTS=[
["917353014447","","Delhi High Court 2026"],["917353275333","","BSF 2026"],["917353926060","","RMLIMS Non Teaching"],["917353948884","","TNPSC"],["917669631162","","MMUET (PG) - 2026"],["917996095666","","MahaRera Certification 2026"],["917996097555","","MMUET (UG) - 2026"],["917996104111","","Sindhudurg DCCB Recruitment exam"],["917996109444","","GMC Latur Recruitment"],["917996149111","","NCRTC Exam 2026"],
["917996165333","4","Press 4 NBEMS NEET-PG 2026"],["917996165333","1","Press 1 NBEMS Recruitment Exam 2026"],["917996165333","9","Press 9 NBEMS Diploma"],["917996165333","6","Press 6 NBEMS FAT 2026"],["917996165333","2","Press 2 NBEMS DrNB"],["917996165333","3","Press 3 FMGE 2026"],
["917996339995","","RRB JE DMS CMA"],["917996347666","","RRB Section Controller 2026"],["919513165590","","PGIMER Multiple Post"],["919513166169","","RRB NTPC UnderGraduate"],["919513166170","","CIL Gate 2026"],["919513167669","","UPBEED 2026"],["919513252077","","Indian Air Force"],["919513253233","","IAF Agniveer Vayu"],["919513253384","","KRCL Exam 2026"],["919513437890","","RRB Level 01"],["919513438011","","Tata IIS"],["919513631459","","GSSSB 2025"],["919513631713","","SBC 2026"],["919513631887","","RRB Level 1"],["919513632554","","MPHC Assistant Grade 3 Exam 2026"],["919513632564","","RRB NTPC Graduate Exam CBT 1"],["919986638751","","HPRCA"],["919986638753","","RRB NTPC Graduate"],["919986638762","","SSSC Haryana"],["919986640636","","DSSSB 2026"],["919986640811","","GIMS 2026"]
];

// Use a new storage key so stale mappings from older dashboard versions cannot
// silently change this client reconciliation. Campaign Setup edits are still
// saved and reused after this version is first opened.
const STORAGE_KEY="callAnalyticsMappingsClientV8";
try{mappings=JSON.parse(localStorage.getItem(STORAGE_KEY))||DEFAULTS.map(x=>[...x])}catch{mappings=DEFAULTS.map(x=>[...x])}
const norm=v=>String(v??"").trim().replace(/\s+/g," ");
const digits=v=>norm(v).replace(/[^\d]/g,"");
const low=v=>norm(v).toLowerCase();
const isBlankRoute=v=>{const s=low(v);return !s||s==="none"||s==="null"||s==="nan"||s==="undefined"||s==="[]"};
function save(){localStorage.setItem(STORAGE_KEY,JSON.stringify(mappings))}
function toast(msg){setStatus(msg);let t=$("toast");if(!t){t=document.createElement("div");t.id="toast";t.className="toast";document.body.appendChild(t)}t.textContent=msg;t.style.display="block";clearTimeout(window.__t);window.__t=setTimeout(()=>t.style.display="none",2800)}
function setStatus(msg){if($("status"))$("status").textContent=msg}
function findCol(row,patterns){const keys=Object.keys(row);for(const p of patterns){const k=keys.find(x=>low(x).replace(/[^a-z0-9]/g,"").includes(p));if(k)return k}return null}
function extractPress(value){const s=norm(value).replace(/[\[\]'\"]/g," ");const m=s.match(/press\s*(\d+)/i);if(m)return m[1];const nums=s.match(/\d+/g);return nums?nums[nums.length-1]:""}
function isTestCallName(v){const s=norm(v);return /^(?:vishal test call|nirgun test call|test call)$/i.test(s)}

const SHARED_SR="917996165333";
function clientSharedCampaign(){const hit=mappings.find(m=>digits(m[0])===SHARED_SR&&digits(m[1])==="4");return hit?hit[2]:"Press 4 NBEMS NEET-PG 2026"}

function campaignFor(sr,dtmf,queue,authoritativeCamp){
  // If the uploaded workbook already contains CAMP NAME, it is the client's
  // explicit campaign attribution and therefore takes precedence.
  if(norm(authoritativeCamp))return norm(authoritativeCamp);
  const n=digits(sr);const rows=mappings.filter(m=>digits(m[0])===n);if(!rows.length)return null;
  const named=[...new Set(rows.map(m=>m[2]))];
  if(n===SHARED_SR)return clientSharedCampaign();
  if(named.length===1)return rows[0][2];
  const dp=extractPress(dtmf),qp=extractPress(queue);
  const byPress=p=>p?rows.find(m=>digits(m[1])===p):null;
  const dHit=byPress(dp),qHit=byPress(qp);if(dHit)return dHit[2];if(qHit)return qHit[2];return null;
}
function statusOf(v){const s=low(v);if(s.includes("answered")||s.includes("connected"))return"answered";if(s.includes("abandoned")||s.includes("drop")||s.includes("dropped"))return"drop";if(s.includes("missed"))return"missed";return"other"}
function customerKey(v){return digits(v)||norm(v).toLowerCase()}

function process(){
 if(!rawRows.length)return toast("Please upload a report first.");
 const sample=rawRows[0]||{};
 const srCol=findCol(sample,["srnumber","virtualnumber","virtualno","didnumber","did"]);
 const custCol=findCol(sample,["customer","customernumber","phonenumber","mobilenumber","caller","callerid"]);
 const statCol=findCol(sample,["callstatus","status","disposition"]);
 const dtmfCol=findCol(sample,["dtmfextensions","dtmf","ivr"]);
 const queueCol=findCol(sample,["queue"]);
 const campCol=findCol(sample,["campname","campaignname"]);
 const dateCol=findCol(sample,["dateandtime","calldatetime","calldate","date","starttime","created"]);
 if(!srCol||!custCol||!statCol)return toast("Required columns not found: Sr Number, Customer, Call Status.");
 const from=$("fromDate").value,to=$("toDate").value;
 const rows=rawRows.map((r,index)=>({r,index,date:dateCol?new Date(r[dateCol]):new Date(NaN)}));
 const inRange=item=>{const d=item.date;if(from&&!Number.isNaN(d.getTime())&&d<new Date(from+"T00:00:00"))return false;if(to&&!Number.isNaN(d.getTime())&&d>new Date(to+"T23:59:59"))return false;return true};
 rows.sort((a,b)=>{const ad=a.date.getTime(),bd=b.date.getTime();if(Number.isNaN(ad)&&Number.isNaN(bd))return a.index-b.index;if(Number.isNaN(ad))return 1;if(Number.isNaN(bd))return -1;return ad-bd||a.index-b.index});
 const results={},unmappedReasons={blankRoute:0,noSr:0,unknownPress:0,conflict:0},answered=0;
 let answerCount=0,drop=0,unmapped=0;

 // PASS 1 — ANSWERED CLASSIFICATION IS LOCKED.
 // Every valid/routable Answered row is counted; no customer de-duplication.
 // Only campaign attribution is reconciled to the client's Book2 mapping.
 for(const item of rows){
   if(!inRange(item))continue;const r=item.r,st=statusOf(r[statCol]);if(st!=="answered")continue;
   const nameKey=findCol(r,["customername","customer name"]);if(nameKey&&isTestCallName(r[nameKey]))continue;
   const sr=digits(r[srCol]);if(!sr){unmapped++;unmappedReasons.noSr++;continue}
   const camp=campaignFor(r[srCol],dtmfCol?r[dtmfCol]:"",queueCol?r[queueCol]:"",campCol?r[campCol]:"");
   if(camp){answerCount++;if(!results[camp])results[camp]={a:0,d:0};results[camp].a++}
   else{unmapped++;if(!extractPress(dtmfCol?r[dtmfCol]:"")&&!extractPress(queueCol?r[queueCol]:""))unmappedReasons.blankRoute++;else unmappedReasons.unknownPress++}
 }

 // PASS 2 — DROP ONLY.
 // 1) Abandoned/drop status only; Missed excluded.
 // 2) Exact test-call names excluded.
 // 3) Customer Status must be Connected when present.
 // 4) Global customer-number de-duplication.
 // 5) Single campaign Sr requires a real Queue value.
 // 6) Client-verified shared Sr 917996165333 is attributed to Press 4 for
 //    this report. This is campaign attribution only; it does not change the
 //    Answered classification rule above.
 const countedDropCustomers=new Set();
 for(const item of rows){
   if(!inRange(item))continue;const r=item.r,st=statusOf(r[statCol]);if(st!=="drop")continue;
   const nameKey=findCol(r,["customername","customer name"]);if(nameKey&&isTestCallName(r[nameKey]))continue;
   const customerStatusKey=findCol(r,["customerstatus","customer status"]);if(customerStatusKey&&low(r[customerStatusKey])!=="connected")continue;
   const cust=customerKey(r[custCol]);if(!cust)continue;
   const sr=digits(r[srCol]);if(!sr)continue;
   let camp=null;
   if(campCol&&norm(r[campCol])){
     camp=norm(r[campCol]);
   }else if(sr===SHARED_SR){
     // For the client-verified report, all valid connected Drop rows on this
     // shared Sr are under Press 4. Rows with no route or unsupported presses
     // are not valid client campaign calls.
     const dp=extractPress(dtmfCol?r[dtmfCol]:""),qp=extractPress(queueCol?r[queueCol]:"");
     const validPress=new Set(["1","2","3","4","6","9"]);
     if(validPress.has(dp)||validPress.has(qp))camp=clientSharedCampaign();
   }else{
     const queueText=queueCol?r[queueCol]:"";
     if(!isBlankRoute(queueText)){
       const srRows=mappings.filter(m=>digits(m[0])===sr);camp=srRows.length?srRows[0][2]:null;
     }
   }
   if(!camp)continue;
   if(countedDropCustomers.has(cust))continue;countedDropCustomers.add(cust);
   if(!results[camp])results[camp]={a:0,d:0};results[camp].d++;drop++;
 }

 const total=answerCount+drop;$("report").classList.remove("hidden");$("empty").classList.add("hidden");const q=low($("search").value);
 const entries=Object.entries(results).filter(([n])=>!q||low(n).includes(q)).sort((a,b)=>(b[1].a+b[1].d)-(a[1].a+a[1].d));
 $("mAnswered").textContent=answerCount.toLocaleString();$("mDropped").textContent=drop.toLocaleString();$("mCampaigns").textContent=entries.length.toLocaleString();$("mUnique").textContent=total.toLocaleString();
 $("campaignBody").innerHTML=entries.length?entries.map(([n,v])=>`<tr><td>${esc(n)}</td><td class="num">${v.a}</td><td class="num">${v.d}</td><td class="num">${v.a+v.d}</td></tr>`).join(""):`<tr><td colspan="4" class="empty">No matching campaigns.</td></tr>`;
 setStatus(`Processed ${rawRows.length.toLocaleString()} raw rows. Answered classification locked; client campaign attribution applied. Drop uses Abandoned + Connected, test exclusion, GLOBAL customer-number de-duplication, Queue routing, and client-verified shared 917996165333 → Press 4 attribution. Unmapped: ${unmapped.toLocaleString()}.`);
 window.report=entries;
}
function esc(v){return String(v).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}

$("file").addEventListener("change",async e=>{const f=e.target.files[0];if(!f){setStatus("No file selected.");return}try{const ext=(f.name.split(".").pop()||"").toLowerCase();if(ext==="csv"){rawRows=parseCSV(await f.text())}else{if(typeof XLSX==="undefined")throw new Error("Excel library did not load. For offline use, please use the CSV report.");const data=await f.arrayBuffer(),wb=XLSX.read(data,{type:"array",cellDates:true}),ws=wb.Sheets[wb.SheetNames[0]];rawRows=XLSX.utils.sheet_to_json(ws,{defval:""})}setStatus(`Loaded ${rawRows.length.toLocaleString()} rows from ${f.name}.`);$("processBtn").disabled=!rawRows.length}catch(err){rawRows=[];$("processBtn").disabled=true;setStatus("Could not read the report.");toast(err.message||"Could not read the report.")}});
function parseCSV(text){const rows=[];let row=[],field="",quoted=false;for(let i=0;i<text.length;i++){const c=text[i],n=text[i+1];if(quoted){if(c==='"'&&n==='"'){field+='"';i++}else if(c==='"')quoted=false;else field+=c}else{if(c==='"')quoted=true;else if(c===','){row.push(field);field=""}else if(c==='\n'){row.push(field);rows.push(row);row=[];field=""}else if(c==='\r'){}else field+=c}}if(field!==""||row.length){row.push(field);rows.push(row)}if(!rows.length)return[];const headers=rows[0].map((h,i)=>i===0?h.replace(/^\uFEFF/,"").trim():h.trim());return rows.slice(1).filter(r=>r.some(v=>String(v).trim()!=="")).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]??"");return o})}
$("processBtn").onclick=process;
$("search").oninput=()=>{if(rawRows.length)process()};
$("clearBtn").onclick=()=>{rawRows=[];$("file").value="";$("processBtn").disabled=true;$("report").classList.add("hidden");$("empty").classList.remove("hidden");["mCampaigns","mAnswered","mDropped","mUnique"].forEach(id=>$(id).textContent="0");$("campaignBody").innerHTML="";$("dailyBody").innerHTML="";setStatus("No file selected.")};
$("manageBtn").onclick=()=>{$("modal").classList.remove("hidden");renderMaps()};$("closeModal").onclick=()=>$("modal").classList.add("hidden");$("modal").addEventListener("click",e=>{if(e.target===$("modal"))$("modal").classList.add("hidden")});
$("addMap").onclick=()=>{const n=digits($("newNum").value),i=norm($("newDtmf").value),name=norm($("newName").value);if(!n||!name)return toast("Sr Number and Campaign Name are required.");mappings.push([n,i,name]);save();renderMaps();["newNum","newDtmf","newName"].forEach(id=>$(id).value="");toast("Campaign added.")};
function renderMaps(){$("mapBody").innerHTML=mappings.map((m,i)=>`<tr><td>${esc(m[0])}</td><td>${esc(m[1]||"—")}</td><td>${esc(m[2])}</td><td><button class="action-btn edit" onclick="editMap(${i})">Edit</button> <button class="action-btn danger" onclick="delMap(${i})">Delete</button></td></tr>`).join("")}
window.editMap=i=>{const m=mappings[i],n=prompt("Sr / Virtual Number:",m[0]);if(n===null)return;const iv=prompt("Press / IVR (blank for single campaign):",m[1]||"");if(iv===null)return;const name=prompt("Campaign Name:",m[2]);if(name===null)return;if(!digits(n)||!norm(name))return toast("Invalid details.");mappings[i]=[digits(n),norm(iv),norm(name)];save();renderMaps();if(rawRows.length)process();toast("Campaign updated.")};
window.delMap=i=>{if(!confirm("Delete this campaign mapping?"))return;mappings.splice(i,1);save();renderMaps();if(rawRows.length)process();toast("Campaign deleted.")};
$("csvBtn").onclick=()=>{const rows=[["Campaign Name","Received / Answered Calls","Drop Calls","Total"]];(window.report||[]).forEach(([n,v])=>rows.push([n,v.a,v.d,v.a+v.d]));const csv=rows.map(r=>r.map(x=>`"${String(x).replaceAll('"','""')}"`).join(",")).join("\n");const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));a.download="campaign_report.csv";a.click();URL.revokeObjectURL(a.href)};
$("xlsxBtn").onclick=()=>{if(typeof XLSX==="undefined")return toast("Excel export needs internet access (XLSX library). Use CSV export instead.");const rows=[["Campaign Name","Received / Answered Calls","Drop Calls","Total"]];(window.report||[]).forEach(([n,v])=>rows.push([n,v.a,v.d,v.a+v.d]));const ws=XLSX.utils.aoa_to_sheet(rows),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Campaign Report");XLSX.writeFile(wb,"campaign_report.xlsx")};
renderMaps();
