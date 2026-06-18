/* =========================================================
   MAC 完整自動化工具 — 程式邏輯 script.js
   分區說明（搭配 修改指南.md 使用）：
   A. 頁面切換          goPage()
   B. 篩選表格產生器     getCols / getMax / updateGen / doDownload
   C. 自動整合處理       上傳事件 / setStep / processFile
   D. 結果預覽與下載     buildPreview / downloadResult / resetAll
   ========================================================= */

/* ══════════ A. 頁面切換 ══════════ */
// 依傳入的 id（home / gen / auto）切換顯示對應頁面與導覽高亮
function goPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));   // 全部頁面先隱藏
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));// 全部導覽取消高亮
  document.getElementById('page-'+id).classList.add('active');                   // 目標頁面顯示
  document.getElementById('nav-'+id).classList.add('active');                    // 目標導覽高亮
  window.scrollTo(0,0);                                                          // 捲回頂端
}

/* ══════════ B. 篩選表格產生器 ══════════ */
// 讀取欄位名稱輸入框 → 用逗號切割、去空白、濾掉空字串 → 回傳陣列
function getCols(){return document.getElementById('colNames').value.split(',').map(s=>s.trim()).filter(Boolean);}
// 讀取最大數字 N，限制 1~99999，無法解析時預設 3000
function getMax(){return Math.max(1,Math.min(99999,parseInt(document.getElementById('maxNum').value)||3000));}

// 依目前設定即時更新統計數字與預覽表格（輸入框一改變就跑）
function updateGen(){
  const n=getMax(),                       // 最大數字 N
        cols=getCols(),                   // 欄位陣列
        dc=Math.max(1,cols.length-1),     // 資料欄數 = 總欄數-1（扣標題欄），至少 1
        dr=n*dc;                          // 資料列數 = N × 資料欄數
  document.getElementById('sRows').textContent=dr.toLocaleString();      // 更新「資料列」
  document.getElementById('sTotal').textContent=(dr+1).toLocaleString(); // 更新「總列數」(+1 標題)
  const show=Math.min(n*dc,16);           // 預覽最多顯示 16 列
  let html='<thead><tr>';
  cols.forEach(c=>{html+=`<th>${c}<span class="flt">▾</span></th>`;});   // 表頭加篩選小圖示
  html+='</tr></thead><tbody>';
  let cnt=0;
  outer:for(let i=1;i<=n;i++){            // 外層 1~N
    for(let d=0;d<dc;d++){                // 內層每個資料欄
      if(cnt>=show)break outer;           // 達到預覽上限就停（一次跳出兩層）
      html+='<tr>';
      cols.forEach((_,ci)=>{
        if(ci===0)html+='<td></td>';                      // 第一欄（標題欄）留空
        else if(ci-1===d)html+=`<td class="v">${i}</td>`; // 對角線位置填入數值 i
        else html+='<td></td>';                           // 其餘留空
      });
      html+='</tr>';cnt++;
    }
  }
  html+='</tbody>';
  document.getElementById('prevTbl').innerHTML=html;       // 寫入預覽表格
  const rem=n*dc-show;                                     // 剩餘未顯示列數
  document.getElementById('tbl-foot').textContent=rem>0?`... 還有 ${rem.toLocaleString()} 列未顯示`:`共 ${(n*dc).toLocaleString()} 列資料`;
}

// 產生並下載篩選範本 Excel
function doDownload(){
  const cols=getCols(),max=getMax(),dc=cols.length-1;       // 欄位 / N / 資料欄數
  const st=document.getElementById('genStatus');            // 狀態訊息元素
  st.className='smsg';st.style.color='var(--muted)';st.textContent='生成中...';
  setTimeout(()=>{                                          // 延遲讓訊息先顯示再運算
    try{
      const wb=XLSX.utils.book_new();                       // 新空白活頁簿
      const data=[cols];                                    // 第一列放欄位標題
      for(let n=1;n<=max;n++){                              // 每個數字
        for(let d=0;d<dc;d++){                              // 每個資料欄各一列
          data.push(cols.map((_,ci)=>ci===0?null:(ci-1===d?n:null))); // 對角線填 n
        }
      }
      const ws=XLSX.utils.aoa_to_sheet(data);               // 二維陣列轉工作表
      ws['!cols']=cols.map((_,i)=>({wch:i===0?16:8}));      // 欄寬：第一欄 16、其餘 8
      ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:0,c:cols.length-1}})}; // 加 AutoFilter
      XLSX.utils.book_append_sheet(wb,ws,'工作表1');         // 加入活頁簿
      XLSX.writeFile(wb,`mac_篩選表格${max}.xlsx`);          // 觸發下載
      st.style.color='var(--green)';
      st.textContent=`✓ 已下載 mac_篩選表格${max}.xlsx`;
    }catch(e){st.className='smsg err';st.textContent='錯誤：'+e.message;}
  },50);
}
// 監聽兩個輸入框：內容改變即時更新預覽
document.getElementById('maxNum').addEventListener('input',updateGen);
document.getElementById('colNames').addEventListener('input',updateGen);
updateGen(); // 載入時先跑一次畫出初始預覽

/* ══════════ C. 自動整合處理 ══════════ */
let resultWB=null,      // 處理完成後的活頁簿物件（供下載）
    baseName='',        // 原始檔名（去副檔名），用於組輸出檔名
    previewRows=[];     // 「工廠使用」工作表完整資料（供預覽）
const PREVIEW_LIMIT=200; // 預覽最多顯示 200 列

// 拖曳進入：加 .drag 視覺回饋
function onDragOver(e){e.preventDefault();document.getElementById('uploadZone').classList.add('drag');}
// 拖曳離開：移除 .drag
function onDragLeave(){document.getElementById('uploadZone').classList.remove('drag');}
// 放下檔案：取第一個檔案處理
function onDrop(e){e.preventDefault();onDragLeave();const f=e.dataTransfer.files[0];if(f)processFile(f);}
// 點擊選檔：取選到的檔案處理
function onFileChange(e){const f=e.target.files[0];if(f)processFile(f);}

// 依步驟 n 更新五個處理膠囊狀態（< n = done，= n = active）
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
  baseName=file.name.replace(/\.[^.]+$/,'');                       // 去副檔名
  document.getElementById('uploadZone').style.display='none';       // 隱藏上傳區
  document.getElementById('processingArea').style.display='block';  // 顯示處理中
  setStep(1);

  const reader=new FileReader();
  reader.onload=function(ev){
    setTimeout(()=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:'array'});        // 解析 Excel

        // ① 找名稱含「号段分配」的工作表
        const srcName=wb.SheetNames.find(n=>n.includes('号段分配'));
        if(!srcName){
          alert('找不到「号段分配」工作表，請確認檔案格式。');
          resetAll(); return;
        }
        const srcWS=wb.Sheets[srcName];
        const srcData=XLSX.utils.sheet_to_json(srcWS,{header:1,defval:null}); // 轉二維陣列

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
                  'pcba sn(按照Thinkstart規則生成)','組裝sn(按照Thinkstart規則生成)']; // 標題列
                const g=macGroups[i-1]||['','','',''];
                return [imei, g[0], g[1], g[2], g[3], '', ''];                       // F、G 留空
              });

              const newWS=XLSX.utils.aoa_to_sheet(newData);

              setStep(5);
              setTimeout(()=>{
                newWS['!cols']=[{wch:18},{wch:16},{wch:16},{wch:16},{wch:16},{wch:28},{wch:28}]; // 欄寬
                XLSX.utils.book_append_sheet(wb,newWS,'工廠使用'); // 加入活頁簿
                resultWB=wb;        // 供下載
                previewRows=newData;// 供預覽

                // 更新結果區 UI
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
  reader.readAsArrayBuffer(file); // 以 ArrayBuffer 讀檔（對應 type:'array'）
}

/* ══════════ D. 結果預覽與下載 ══════════ */
// 把 previewRows 畫成 HTML 表格（最多 PREVIEW_LIMIT 列）
function buildPreview(){
  const limit=Math.min(previewRows.length,PREVIEW_LIMIT);
  let html='';
  for(let i=0;i<limit;i++){
    const r=previewRows[i];
    const isHdr=i===0;
    html+='<tr>';
    html+=`<td class="cy ${isHdr?'':'cv'}">${r[0]||''}</td>`; // A IMEI 黃底
    html+=`<td class="${r[1]?'cv':'ce'}">${r[1]||''}</td>`;   // B mac1
    html+=`<td class="${r[2]?'cv':'ce'}">${r[2]||''}</td>`;   // C mac2
    html+=`<td class="${r[3]?'cv':'ce'}">${r[3]||''}</td>`;   // D mac3
    html+=`<td class="${r[4]?'cv':'ce'}">${r[4]||''}</td>`;   // E mac4
    html+=`<td class="ce">${r[5]||''}</td>`;                  // F pcba sn
    html+=`<td class="ce">${r[6]||''}</td>`;                  // G 組裝sn
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

// 清空所有狀態，回到最初上傳畫面
function resetAll(){
  resultWB=null; previewRows=[];
  document.getElementById('resultArea').style.display='none';
  document.getElementById('processingArea').style.display='none';
  document.getElementById('uploadZone').style.display='block';
  document.getElementById('fileInput').value=''; // 清空已選檔（可重選同檔）
  document.getElementById('dlMsg').textContent='';
}
