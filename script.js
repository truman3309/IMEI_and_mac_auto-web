/* =========================================================
   號段分配工具集 — 合併程式邏輯
   區塊：A. 頁面切換
        B. 功能① 欄位堆疊（前綴 st = stack）
        C. 功能② MAC 自動整合（沿用原 IMEI/MAC 工具）
   兩功能各自獨立，ID 已分開（堆疊頁一律 st 開頭）避免衝突。
   ========================================================= */

/* ══════════ A. 頁面切換 ══════════ */
function goPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  document.getElementById('nav-'+id).classList.add('active');
  window.scrollTo(0,0);
}

/* ══════════ B. 功能① 欄位堆疊 ══════════ */
(function(){
  const $ = id => document.getElementById(id);
  const drop = $('stDrop'), fileInput = $('stFile'),
        result = $('stResult'), errBox = $('stError');
  let outWb = null, outName = 'merged.xlsx';

  // 把任意儲存格值轉成乾淨字串，避免科學記號
  function S(v){
    if(v===null||v===undefined) return '';
    if(typeof v==='number') return Number.isInteger(v) ? v.toString() : String(v);
    return String(v).trim();
  }
  function showError(msg){ errBox.textContent=msg; errBox.classList.add('show'); result.classList.remove('show'); }
  function clearError(){ errBox.classList.remove('show'); }

  function handleFile(file){
    clearError();
    if(!file) return;
    const ok = /\.(xlsx|xlsm|csv)$/i.test(file.name);
    if(!ok){ showError('請選擇 .xlsx、.xlsm 或 .csv 檔案。'); return; }
    const reader = new FileReader();
    reader.onload = e => {
      try { process(e.target.result, file.name); }
      catch(err){ console.error(err); showError('讀取失敗：' + (err.message || '檔案格式無法解析')); }
    };
    reader.onerror = () => showError('讀取檔案時發生錯誤，請再試一次。');
    reader.readAsArrayBuffer(file);
  }

  function process(buffer, originalName){
    const wb = XLSX.read(new Uint8Array(buffer), {type:'array'});
    if(!wb.SheetNames.length){ showError('檔案裡找不到任何工作表。'); return; }
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null, blankrows:false});
    if(aoa.length < 1){ showError('工作表是空的，沒有資料可以處理。'); return; }

    const header = aoa[0] || [];
    const body = aoa.slice(1);

    // 各欄獨立收集非空白值
    const colA=[], colB=[], colC=[], colD=[];
    for(const row of body){
      const a=S(row && row[0]), b=S(row && row[1]), c=S(row && row[2]), d=S(row && row[3]);
      if(a!=='') colA.push(a);
      if(b!=='') colB.push(b);
      if(c!=='') colC.push(c);
      if(d!=='') colD.push(d);
    }
    if(colC.length===0 && colD.length===0){
      showError('C 欄和 D 欄都沒有資料，沒有東西可以接到底下。請確認檔案是 A／B 與 C／D 兩組並排的格式。');
      return;
    }

    // C 接 A 底下、D 接 B 底下
    const imei = colA.concat(colC);
    const mac  = colB.concat(colD);
    const headA = S(header[0]) || 'IMEI';
    const headB = S(header[1]) || 'MAC';

    // 組輸出 AOA
    const out = [[headA, headB]];
    const n = Math.max(imei.length, mac.length);
    for(let i=0;i<n;i++) out.push([ imei[i]||'', mac[i]||'' ]);

    // 全部以文字型態寫入，避免 Excel 轉成科學記號
    const outWs = XLSX.utils.aoa_to_sheet(out);
    const range = XLSX.utils.decode_range(outWs['!ref']);
    for(let R=range.s.r; R<=range.e.r; R++){
      for(let C=range.s.c; C<=range.e.c; C++){
        const addr = XLSX.utils.encode_cell({r:R,c:C});
        const cell = outWs[addr];
        if(cell && cell.v!==''){ cell.t='s'; cell.v=String(cell.v); cell.z='@'; }
      }
    }
    outWs['!cols'] = [{wch:20},{wch:18}];
    outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, outWs, sheetName);
    outName = originalName.replace(/\.(xlsx|xlsm|csv)$/i,'') + '_合併完成.xlsx';

    render({ originalName, colA:colA.length, colB:colB.length, colC:colC.length, colD:colD.length,
             imei, mac });
  }

  function render(d){
    clearError();
    $('stFname').textContent = d.originalName;
    $('stImeiTotal').innerHTML = d.imei.length.toLocaleString() + ' <small>筆</small>';
    $('stMacTotal').innerHTML  = d.mac.length.toLocaleString()  + ' <small>筆</small>';
    $('stImeiBreak').innerHTML = `原 A 欄 <b>${d.colA.toLocaleString()}</b> ＋ C 欄移入 <b>${d.colC.toLocaleString()}</b>`;
    $('stMacBreak').innerHTML  = `原 B 欄 <b>${d.colB.toLocaleString()}</b> ＋ D 欄移入 <b>${d.colD.toLocaleString()}</b>`;

    // 預覽：接合點前後各幾筆（以 A 欄筆數為接合點）
    const tb = $('stPreviewBody'); tb.innerHTML = '';
    const join = d.colA;
    const total = Math.max(d.imei.length, d.mac.length);
    const head = [0,1,2];
    const around = [join-1, join, join+1].filter(i=>i>=0 && i<total);
    const tail = [total-2, total-1].filter(i=>i>=0);
    const set = [...new Set([...head, ...around, ...tail])].filter(i=>i>=0 && i<total).sort((a,b)=>a-b);

    let prev = -1;
    for(const i of set){
      if(prev!==-1 && i>prev+1){
        const tr=document.createElement('tr'); tr.className='gap';
        tr.innerHTML='<td colspan="3">· · ·</td>'; tb.appendChild(tr);
      }
      const fromMove = i>=join;
      const tr=document.createElement('tr');
      if(fromMove) tr.className='from-move';
      tr.innerHTML = `<td class="idx">${i+2}</td><td>${d.imei[i]||''}</td><td>${d.mac[i]||''}</td>`;
      tb.appendChild(tr);
      prev=i;
    }
    result.classList.add('show');
    result.scrollIntoView({behavior:'smooth', block:'nearest'});
  }

  // 互動綁定
  $('stDownload').addEventListener('click', () => { if(outWb) XLSX.writeFile(outWb, outName); });
  $('stReset').addEventListener('click', () => {
    outWb=null; fileInput.value=''; result.classList.remove('show'); clearError();
    drop.scrollIntoView({behavior:'smooth', block:'center'});
  });
  drop.addEventListener('click', ()=>fileInput.click());
  drop.addEventListener('keydown', e=>{ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); fileInput.click(); }});
  fileInput.addEventListener('change', e=>handleFile(e.target.files[0]));
  ['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev, e=>{ e.preventDefault(); drop.classList.add('dragging'); }));
  ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev, e=>{ e.preventDefault(); drop.classList.remove('dragging'); }));
  drop.addEventListener('drop', e=>{ const f=e.dataTransfer.files[0]; handleFile(f); });
})();

/* ══════════ C. 功能② MAC 自動整合 ══════════ */
let resultWB=null,      // 處理完成後的活頁簿（供下載）
    baseName='',        // 原始檔名（去副檔名）
    previewRows=[];     // 「工廠使用」完整資料（供預覽）
const PREVIEW_LIMIT=200;

function onDragOver(e){e.preventDefault();document.getElementById('uploadZone').classList.add('drag');}
function onDragLeave(){document.getElementById('uploadZone').classList.remove('drag');}
function onDrop(e){e.preventDefault();onDragLeave();const f=e.dataTransfer.files[0];if(f)processFile(f);}
function onFileChange(e){const f=e.target.files[0];if(f)processFile(f);}

// 依步驟 n 更新五個處理膠囊狀態
function setStep(n){
  for(let i=1;i<=5;i++){
    const el=document.getElementById('ps'+i);
    if(i<n)el.className='p-step done';
    else if(i===n)el.className='p-step active';
    else el.className='p-step';
  }
}

// 核心流程：讀取上傳 Excel → 整理資料 → 建立「工廠使用」工作表
function processFile(file){
  baseName=file.name.replace(/\.[^.]+$/,'');
  document.getElementById('uploadZone').style.display='none';
  document.getElementById('processingArea').style.display='block';
  document.getElementById('resultArea').style.display='none';
  setStep(1);

  const reader=new FileReader();
  reader.onload=function(ev){
    setTimeout(()=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:'array'});

        // ① 找名稱含「号段分配」的工作表
        const srcName=wb.SheetNames.find(n=>n.includes('号段分配'));
        if(!srcName){
          alert('找不到「号段分配」工作表，請確認檔案格式。');
          resetAll(); return;
        }
        const srcWS=wb.Sheets[srcName];
        const srcData=XLSX.utils.sheet_to_json(srcWS,{header:1,defval:null});
        const imeiData=srcData.map(r=>r[0]!=null?String(r[0]):'');                  // A欄 IMEI（含標題）
        const macRaw=srcData.slice(1).map(r=>r[1]!=null?String(r[1]):'').filter(v=>v); // B欄 MAC（去標題去空值）

        setStep(2);
        setTimeout(()=>{
          // ② MAC 每 4 筆一組（橫向 mac1~mac4）
          const numCols=4;
          const totalGroups=Math.ceil(macRaw.length/numCols);
          const macGroups=[];
          for(let i=0;i<totalGroups;i++){
            const g=i*numCols;
            macGroups.push([macRaw[g]||'', macRaw[g+1]||'', macRaw[g+2]||'', macRaw[g+3]||'']);
          }

          setStep(3);
          setTimeout(()=>{
            // ③ 移除舊「工廠使用」工作表（若存在）
            if(wb.SheetNames.includes('工廠使用')){
              const idx=wb.SheetNames.indexOf('工廠使用');
              wb.SheetNames.splice(idx,1);
              delete wb.Sheets['工廠使用'];
            }

            setStep(4);
            setTimeout(()=>{
              // ④ 組「工廠使用」二維資料：A=IMEI,B~E=mac1~4,F=pcba sn,G=組裝sn
              const newData=imeiData.map((imei,i)=>{
                if(i===0) return ['IMEI','mac1','mac2','mac3','mac4',
                  'pcba sn(按照Thinkstart規則生成)','組裝sn(按照Thinkstart規則生成)'];
                const g=macGroups[i-1]||['','','',''];
                return [imei, g[0], g[1], g[2], g[3], '', ''];
              });
              const newWS=XLSX.utils.aoa_to_sheet(newData);

              setStep(5);
              setTimeout(()=>{
                newWS['!cols']=[{wch:18},{wch:16},{wch:16},{wch:16},{wch:16},{wch:28},{wch:28}];
                XLSX.utils.book_append_sheet(wb,newWS,'工廠使用');
                resultWB=wb;
                previewRows=newData;

                document.getElementById('processingArea').style.display='none';
                document.getElementById('fName').textContent=file.name;
                document.getElementById('fInfo').textContent=`${wb.SheetNames.length} 個工作表`;
                document.getElementById('sumImei').textContent=`${(imeiData.length-1).toLocaleString()} 筆`;
                document.getElementById('sumMac').textContent=`${totalGroups.toLocaleString()} 組`;
                document.getElementById('previewCount').textContent=`${newData.length.toLocaleString()} 列`;

                buildPreview();
                document.getElementById('resultArea').style.display='block';
              },100);
            },100);
          },100);
        },100);
      }catch(err){
        alert('處理失敗：'+err.message);
        resetAll();
      }
    },200);
  };
  reader.readAsArrayBuffer(file);
}

// 把 previewRows 畫成 HTML 表格（最多 PREVIEW_LIMIT 列）
function buildPreview(){
  const limit=Math.min(previewRows.length,PREVIEW_LIMIT);
  let html='';
  for(let i=0;i<limit;i++){
    const r=previewRows[i];
    const isHdr=i===0;
    html+='<tr>';
    html+=`<td class="cy ${isHdr?'':'cv'}">${r[0]||''}</td>`;
    html+=`<td class="${r[1]?'cv':'ce'}">${r[1]||''}</td>`;
    html+=`<td class="${r[2]?'cv':'ce'}">${r[2]||''}</td>`;
    html+=`<td class="${r[3]?'cv':'ce'}">${r[3]||''}</td>`;
    html+=`<td class="${r[4]?'cv':'ce'}">${r[4]||''}</td>`;
    html+=`<td class="ce">${r[5]||''}</td>`;
    html+=`<td class="ce">${r[6]||''}</td>`;
    html+='</tr>';
  }
  document.getElementById('previewBody').innerHTML=html;
  const rem=previewRows.length-limit;
  document.getElementById('previewFoot').textContent=rem>0
    ?`顯示前 ${limit.toLocaleString()} / ${previewRows.length.toLocaleString()} 列`
    :`共 ${previewRows.length.toLocaleString()} 列`;
}

// 下載處理完成的 Excel
function downloadResult(){
  if(!resultWB) return;
  const st=document.getElementById('dlMsg');
  st.className='smsg';st.textContent='生成中...';
  setTimeout(()=>{
    try{
      XLSX.writeFile(resultWB,`${baseName}_明細表.xlsx`);
      st.textContent='✓ 已下載完成';
    }catch(e){st.className='smsg err';st.textContent='錯誤：'+e.message;}
  },50);
}

// 清空狀態，回到上傳畫面
function resetAll(){
  resultWB=null; previewRows=[];
  document.getElementById('resultArea').style.display='none';
  document.getElementById('processingArea').style.display='none';
  document.getElementById('uploadZone').style.display='block';
  document.getElementById('fileInput').value='';
  document.getElementById('dlMsg').textContent='';
}
