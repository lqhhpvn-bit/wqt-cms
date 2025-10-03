/***** model.gs — helpers đọc/ghi theo schema (VI) *****/

function _ss_(){ return SpreadsheetApp.getActive(); }
function _getSheet_(name){
  const sh = _ss_().getSheetByName(name);
  if (!sh) throw new Error('Không tìm thấy sheet: ' + name);
  return sh;
}
function _headers_(sh){
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
}
function _readAsObjects_(sheetName){
  const sh = _getSheet_(sheetName);
  const headers = _headers_(sh);
  if (sh.getLastRow()<2) return [];
  const vals = sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues();
  return vals.map(row=>{
    const o={}; headers.forEach((h,i)=> o[h] = row[i]==='' ? '' : row[i]); return o;
  });
}
function _findRowById_(sheetName, idCol, idVal){
  const sh = _getSheet_(sheetName);
  const headers = _headers_(sh);
  const idIdx = headers.indexOf(idCol);
  if (idIdx<0) throw new Error('Thiếu cột ID: ' + idCol);
  const n = Math.max(0, sh.getLastRow()-1);
  if (n===0) return {row:-1, headers};
  const vals = sh.getRange(2,1,n,headers.length).getValues();
  for (let i=0;i<n;i++){
    if (String(vals[i][idIdx])===String(idVal)) return {row:i+2, headers};
  }
  return {row:-1, headers};
}
function _appendObject_(sheetName, obj){
  const sh = _getSheet_(sheetName);
  const headers = _headers_(sh);
  const arr = headers.map(h => obj[h]!==undefined ? obj[h] : '');
  sh.appendRow(arr);
}
function _writeRowByIndex_(sh, row1, headers, obj){
  const arr = headers.map(h => obj[h]!==undefined ? obj[h] : '');
  sh.getRange(row1,1,1,headers.length).setValues([arr]);
}
function _filterSearchPaginate_(rows, query, page, pageSize, sortBy, sortDir){
  query = String(query||'').trim().toLowerCase();
  let out = rows;
  if (query) out = out.filter(o => JSON.stringify(o).toLowerCase().includes(query));
  if (sortBy){
    out.sort((a,b)=>{
      const A=(a[sortBy]??'').toString(), B=(b[sortBy]??'').toString();
      return sortDir==='desc' ? B.localeCompare(A) : A.localeCompare(B);
    });
  }
  const total = out.length;
  const from = Math.max((page-1)*pageSize,0);
  const to = Math.min(from+pageSize,total);
  return { total, items: out.slice(from,to) };
}
function _genId_(prefix){
  const d = new Date(), p=n=>String(n).padStart(2,'0');
  const base = d.getFullYear().toString().slice(2)+p(d.getMonth()+1)+p(d.getDate())+p(d.getHours())+p(d.getMinutes())+p(d.getSeconds());
  const rnd = Math.floor(Math.random()*1000).toString().padStart(3,'0');
  return prefix + '-' + base + rnd;
}
