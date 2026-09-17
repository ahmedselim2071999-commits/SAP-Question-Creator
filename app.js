const $=id=>document.getElementById(id);
let questions=[], templateBuffer=null;
const codes={"Pre Assessment":"PA","Knowledge Check":"KC","Post Assessment":"POST"};
const aliases={"SINGLE":"MULTI_CHOICE_SINGLE_ANSWER","SINGLE_ANSWER":"MULTI_CHOICE_SINGLE_ANSWER","MULTIPLE":"MULTI_CHOICE_MULTIPLE_ANSWER","MULTIPLE_ANSWER":"MULTI_CHOICE_MULTIPLE_ANSWER","TRUEFALSE":"TRUE_FALSE","ORDER":"ORDERING","FILL_BLANK":"FILL_IN_THE_BLANK","FILL_IN_BLANK":"FILL_IN_THE_BLANK"};
const supported=new Set(["MULTI_CHOICE_SINGLE_ANSWER","MULTI_CHOICE_MULTIPLE_ANSWER","TRUE_FALSE","ORDERING","FILL_IN_THE_BLANK"]);
function status(t,k=""){let s=$("status");s.textContent=t;s.className="status "+k}
function safe(s){return s.trim().replace(/[^A-Za-z0-9_.-]+/g,"_").replace(/^_+|_+$/g,"")}
function typeOf(s){let v=s.trim().toUpperCase().replace(/[ -]/g,"_");v=aliases[v]||v;if(!supported.has(v))throw Error("Unsupported question type: "+v);return v}
function idFor(n){let c=safe($("course").value)||"COURSE",a=codes[$("assessment").value]||"ASSESSMENT";return `${c}_${a}_${String(n).padStart(3,"0")}`}
function updateId(){$("idExample").textContent=idFor(Number($("startNumber").value)||1)}
["course","assessment","startNumber"].forEach(x=>$(x).addEventListener("input",updateId));$("assessment").addEventListener("change",updateId);
$("questionsFile").addEventListener("change",e=>$("questionsName").textContent=e.target.files[0]?.name||"No file selected");
$("templateFile").addEventListener("change",async e=>{let f=e.target.files[0];$("templateName").textContent=f?.name||"No file selected";templateBuffer=f?await f.arrayBuffer():null});
async function getLines(f){if(f.name.toLowerCase().endsWith(".docx")){let r=await mammoth.extractRawText({arrayBuffer:await f.arrayBuffer()});return r.value.split(/\r?\n/)}return (await f.text()).split(/\r?\n/)}
function parse(lines){
 let out=[],cur=null,answer=null,typ=null;
 const ar=/^(?:ANSWER|CORRECT ANSWER|ANS)\s*:\s*(.+)$/i,tr=/^(?:TYPE|QUESTION TYPE)\s*:\s*(.+)$/i,cr=/^([A-Za-z])[\.\)\:\-]\s*(.+)$/;
 function flush(){if(!cur)return;let text=cur.text.trim(),choices=cur.choices.slice(),t=typeOf(typ||$("defaultType").value),correct=new Set();
  if(t==="FILL_IN_THE_BLANK"){if(!answer)throw Error("Missing ANSWER for: "+text);choices=[{text:answer.trim()}];correct.add(0)}
  else{if(!choices.length)throw Error("No choices found for: "+text);if(!answer)throw Error("Missing ANSWER for: "+text);
   for(let x of answer.split(/[,;|]/).map(v=>v.trim()).filter(Boolean)){if(/^[A-Za-z]$/.test(x)){let i=x.toUpperCase().charCodeAt(0)-65;if(i>=choices.length)throw Error("Answer "+x+" is outside choices: "+text);correct.add(i)}else{let i=choices.findIndex(c=>c.text.toLowerCase()===x.toLowerCase());if(i<0)throw Error("Could not match answer '"+x+"': "+text);correct.add(i)}}
   if(t==="MULTI_CHOICE_SINGLE_ANSWER"&&correct.size!==1)throw Error("Single-answer question needs exactly one correct answer: "+text);
   if(t==="MULTI_CHOICE_MULTIPLE_ANSWER"&&!correct.size)throw Error("Multiple-answer question needs a correct answer: "+text)}
  out.push({text,choices,correct,type:t});cur=null;answer=null;typ=null}
 for(let raw of lines){let l=raw.trim();if(!l){if(cur&&answer!==null)flush();continue}
  let m=l.match(ar);if(m){if(!cur)throw Error("ANSWER appeared before a question.");answer=m[1].trim();continue}
  m=l.match(tr);if(m){typ=m[1].trim();continue}
  m=l.match(cr);if(m){if(!cur)throw Error("Choice appeared before a question.");cur.choices.push({text:m[2].trim()});continue}
  if(cur&&answer!==null)flush();if(!cur)cur={text:l,choices:[]};else cur.text+=" "+l}
 flush();return out
}
function esc(s){return String(s).replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",""":"&quot;","'":"&#039;"}[m]))}
function render(){let b=$("tbody");b.innerHTML="";let start=Number($("startNumber").value)||1;questions.forEach((q,i)=>{let r=document.createElement("tr"),cor=[...q.correct].map(n=>String.fromCharCode(65+n)).join(", ")||"—";r.innerHTML=`<td>${i+1}</td><td>${esc(q.text)}</td><td><code>${q.type}</code></td><td>${q.choices.length}</td><td>${cor}</td><td><code>${idFor(start+i)}</code></td><td><span class="good">✓ Valid</span></td>`;b.appendChild(r)});
 $("summary").classList.remove("hidden");$("summary").innerHTML=`<span class="pill">${questions.length} questions</span><span class="pill">${questions.reduce((a,q)=>a+q.choices.length,0)} choice rows</span><span class="pill">0 errors</span>`;$("generateBtn").disabled=!templateBuffer||!questions.length}
$("parseBtn").onclick=async()=>{try{let f=$("questionsFile").files[0];if(!f)return status("Please select a question file.","error");status("Parsing questions…");questions=parse(await getLines(f));render();status(`Parsed ${questions.length} question(s). Review before export.`,"ok")}catch(e){questions=[];$("generateBtn").disabled=true;status(e.message,"error")}}
function copyStyle(to,from){to.height=from.height;for(let i=1;i<=from.cellCount;i++){let a=from.getCell(i),b=to.getCell(i);try{b.style={...a.style}}catch(_){}try{b.alignment={...a.alignment}}catch(_){}try{b.border={...a.border}}catch(_){}try{b.fill={...a.fill}}catch(_){}try{b.font={...a.font}}catch(_){}b.numFmt=a.numFmt}}
function cell(ws,row,heads,key,val){ws.getCell(row,heads[key]).value=val}
$("generateBtn").onclick=async()=>{try{status("Generating SAP workbook in your browser…");let wb=new ExcelJS.Workbook();await wb.xlsx.load(templateBuffer);let ws=wb.getWorksheet("Question");if(!ws)throw Error("Template must contain a 'Question' sheet.");
 let heads={};ws.getRow(1).eachCell((c,n)=>heads[String(c.value||"").trim()]=n);let req=["Question ID (*required)","Locale ID (*required)","Domain ID (*required)","Variant Number (*required)","Question Type (*required)","Question Text (*required)","Answer Choice Number (*required)","Answer Choice Value","Is Correct (*required)"];let miss=req.filter(x=>!heads[x]);if(miss.length)throw Error("Template is missing columns: "+miss.join(", "));
 let source=ws.getRow(2);if(ws.rowCount>1)ws.spliceRows(2,ws.rowCount-1);let row=2,start=Number($("startNumber").value)||1;
 questions.forEach((q,qi)=>{let choices=q.choices;if(q.type==="TRUE_FALSE"&&!choices.length)choices=[{text:"True"},{text:"False"}];choices.forEach((ch,ci)=>{if(row>2)copyStyle(ws.getRow(row),source);let v={"Question ID (*required)":idFor(start+qi),"Locale ID (*required)":$("locale").value.trim()||"en_US","Domain ID (*required)":$("domain").value.trim()||"PUBLIC","Variant Number (*required)":1,"Question Type (*required)":q.type,"Question Text (*required)":q.text,"Answer Choice Number (*required)":ci+1,"Answer Choice Value":ch.text,"Is Correct (*required)":q.correct.has(ci)?"Y":"N"};if(heads["Point Value"])v["Point Value"]=1;if(heads["Available for Exams"])v["Available for Exams"]="Y";Object.entries(v).forEach(([k,x])=>cell(ws,row,heads,k,x));row++})});
 ["Feedback","Objectives"].forEach(n=>{let s=wb.getWorksheet(n);if(s&&s.rowCount>1)s.spliceRows(2,s.rowCount-1)});
 let buf=await wb.xlsx.writeBuffer(),blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),url=URL.createObjectURL(blob),a=document.createElement("a"),course=safe($("course").value)||"SAP_LMS",code=codes[$("assessment").value]||"ASSESSMENT";a.href=url;a.download=`${course}_${code}_SAP_Bulk.xlsx`;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);status(`Done — ${questions.length} questions and ${row-2} choice rows generated.`,"ok")
}catch(e){console.error(e);status(e.message||"Generation failed.","error")}}
$("demoBtn").onclick=()=>{let text=`Which formula represents Einstein's mass-energy equivalence?\n\nA. F = ma\nB. E = mc^2\nC. PV = nRT\nD. V = IR\n\nANSWER: B\n\nWhich law relates voltage, current, and resistance?\n\nA. Newton's Law\nB. Ohm's Law\nC. Hubble's Law\nD. Boyle's Law\n\nANSWER: B`;let dt=new DataTransfer();dt.items.add(new File([text],"demo.aiken",{type:"text/plain"}));$("questionsFile").files=dt.files;$("questionsName").textContent="demo.aiken";$("course").value="26Q3_vid_c_SOWA2_JB";updateId();status("Demo loaded. Upload the SAP template, then parse.","ok")}
$("clearBtn").onclick=()=>{location.reload()}
updateId();