const COLS=5, START_ROWS=10;
let rows=[], activePhotoRow=null, objectUrls=new Set();

const tbody=document.querySelector("#sheet tbody");
const status=document.querySelector("#status");

function makeRow(data={}){return {values:Array.from({length:COLS},(_,i)=>data.values?.[i]??""), photo:data.photo||null};}
function init(){rows=Array.from({length:START_ROWS},()=>makeRow()); render(); setStatus("New sheet");}
function setStatus(s){status.textContent=s; setTimeout(()=>{if(status.textContent===s)status.textContent=""},2500)}

function render(){
  objectUrls.forEach(u=>URL.revokeObjectURL(u)); objectUrls.clear();
  tbody.innerHTML="";
  rows.forEach((r,i)=>{
    const tr=document.createElement("tr");
    tr.dataset.i=i;
    const c=document.createElement("td"); c.className="check"; c.innerHTML='<input type="checkbox" class="rowcheck">'; tr.appendChild(c);
    const n=document.createElement("td"); n.textContent=i+1; tr.appendChild(n);
    for(let j=0;j<COLS;j++){
      const td=document.createElement("td");
      const inp=document.createElement("input"); inp.className="cell"; inp.value=r.values[j]||"";
      inp.addEventListener("input",e=>r.values[j]=e.target.value);
      td.appendChild(inp); tr.appendChild(td);
    }
    const p=document.createElement("td"); p.className="photo-cell"; renderPhoto(p,r,i); tr.appendChild(p);
    tbody.appendChild(tr);
  });
}
function renderPhoto(td,r,i){
  if(r.photo){
    const img=document.createElement("img"); img.className="photo-thumb";
    img.src=typeof r.photo==="string" ? r.photo : URL.createObjectURL(r.photo);
    objectUrls.add(img.src);
    td.appendChild(img);
    const actions=document.createElement("div"); actions.className="photo-actions";
    const replace=document.createElement("button"); replace.textContent="Replace"; replace.onclick=()=>openPhotoMenu(i);
    const remove=document.createElement("button"); remove.textContent="Remove"; remove.onclick=()=>{r.photo=null;render()};
    actions.append(replace,remove); td.appendChild(actions);
  }else{
    const b=document.createElement("button"); b.textContent="📷 Browse"; b.onclick=()=>openPhotoMenu(i); td.appendChild(b);
    const d=document.createElement("div"); d.className="empty"; d.textContent="No photo"; td.appendChild(d);
  }
}
function openPhotoMenu(i){activePhotoRow=i;document.querySelector("#photoMenu").classList.remove("hidden")}
function closePhotoMenu(){document.querySelector("#photoMenu").classList.add("hidden");activePhotoRow=null}
document.querySelector("#cancelPhoto").onclick=closePhotoMenu;
document.querySelector("#cameraBtn").onclick=()=>{closePhotoMenu();document.querySelector("#cameraInput").click()};
document.querySelector("#galleryBtn").onclick=()=>{closePhotoMenu();document.querySelector("#galleryInput").click()};
["cameraInput","galleryInput"].forEach(id=>document.getElementById(id).addEventListener("change",e=>{
  const f=e.target.files[0]; if(!f)return;
  if(activePhotoRow!==null) rows[activePhotoRow].photo=f;
  render(); setStatus("Photo added"); e.target.value="";
}));

document.querySelector("#addRowBtn").onclick=()=>{rows.push(makeRow());render();setStatus("Row added")};
document.querySelector("#deleteRowsBtn").onclick=()=>{
  const checks=[...document.querySelectorAll(".rowcheck")];
  const remove=new Set(checks.map((c,i)=>c.checked?i:-1).filter(i=>i>=0));
  if(!remove.size){setStatus("Select rows first");return}
  rows=rows.filter((_,i)=>!remove.has(i)); if(!rows.length)rows.push(makeRow()); render(); setStatus("Rows deleted");
};
document.querySelector("#selectAll").onchange=e=>document.querySelectorAll(".rowcheck").forEach(c=>c.checked=e.target.checked);
document.querySelector("#newBtn").onclick=()=>{if(confirm("Clear the current sheet and start a new one?"))init()};

async function downloadExcel(){
  const wb=new ExcelJS.Workbook();
  const ws=wb.addWorksheet("Sheet1");
  ws.columns=[
    {header:"Column 1",key:"c1",width:22},{header:"Column 2",key:"c2",width:22},
    {header:"Column 3",key:"c3",width:22},{header:"Column 4",key:"c4",width:22},
    {header:"Column 5",key:"c5",width:22},{header:"Photo",key:"photo",width:24}
  ];
  ws.getRow(1).font={bold:true}; ws.getRow(1).height=22;
  for(let i=0;i<rows.length;i++){
    const r=rows[i], er=ws.addRow(r.values);
    er.height=r.photo?85:20;
    if(r.photo){
      const buf=await blobToBuffer(r.photo);
      const ext=(r.photo.name||"photo.jpg").toLowerCase().endsWith(".png")?"png":"jpeg";
      const imageId=wb.addImage({buffer:buf,extension:ext});
      ws.addImage(imageId,{tl:{col:5.15,row:i+1.15},ext:{width:120,height:75}});
    }
  }
  ws.views=[{state:"frozen",ySplit:1}];
  const buffer=await wb.xlsx.writeBuffer();
  const blob=new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="photo_excel_sheet.xlsx";a.click();
  setStatus("Excel downloaded");
}
function blobToBuffer(blobOrData){
  if(blobOrData instanceof Blob)return blobOrData.arrayBuffer();
  if(typeof blobOrData==="string")return fetch(blobOrData).then(r=>r.arrayBuffer());
}
document.querySelector("#downloadBtn").onclick=downloadExcel;

document.querySelector("#openInput").addEventListener("change",async e=>{
  const file=e.target.files[0]; if(!file)return;
  try{
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
    const ws=wb.worksheets[0]; const out=[];
    ws.eachRow((row,idx)=>{
      if(idx===1)return;
      const vals=[]; for(let j=1;j<=COLS;j++)vals.push(row.getCell(j).text||"");
      if(vals.some(v=>v.trim()) || idx<=ws.rowCount)out.push(makeRow({values:vals}));
    });
    // Recover embedded images and associate by row based on image anchor.
    const imgs=ws.getImages();
    for(const pic of imgs){
      const row=Math.floor(pic.range.tl.row-1);
      if(row>=0 && row<out.length){
        const image=wb.getImage(pic.imageId);
        if(image && image.buffer){
          const mime=image.extension==="png"?"image/png":"image/jpeg";
          out[row].photo=new Blob([image.buffer],{type:mime});
        }
      }
    }
    rows=out.length?out:Array.from({length:START_ROWS},()=>makeRow());
    render(); setStatus(`Opened ${file.name}`);
  }catch(err){console.error(err);alert("Could not open this Excel file. Please use an .xlsx file created by this app.");}
  e.target.value="";
});
init();
