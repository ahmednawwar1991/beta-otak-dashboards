/* Browser-side extraction — mirrors extract.py. Needs SheetJS (XLSX) loaded first.
   Exposes window.XLSXExtract.fromWorkbooks([wb, ...]) -> {contracts, orders, msr}. */
(function(){
  function pad2(x){return String(x).padStart(2,'0');}
  function fmtDate(d){return d.getUTCFullYear()+'-'+pad2(d.getUTCMonth()+1)+'-'+pad2(d.getUTCDate());}
  function clean(v){
    if(v==null) return null;
    if(v instanceof Date && !isNaN(v)) return fmtDate(v);
    if(typeof v==='string'){const s=v.trim(); if(s===''||s[0]==='#') return null; return s;}
    if(typeof v==='number'){return Number.isInteger(v)?v:v;}
    return v;
  }
  function num(v){
    v=clean(v);
    if(typeof v==='number') return v;
    if(typeof v==='string'){const s=v.replace(/,/g,'').replace(/%/g,'').trim(); const n=parseFloat(s); return isNaN(n)?null:n;}
    return null;
  }
  var PROJ={'code 8':'Code 8','Code 8':'Code 8','demerdash hospital':'Demerdash Hospital',
            'جوهرة العاصمه':'Gawharet El Asema','نادى الرمايه الدولى':"Int'l Shooting Club"};
  function projn(p){p=(p==null?'':String(p)).trim(); return PROJ[p]||p;}
  var CT={'مصنعيات':'Workmanship','عمالة':'Labor','توريدات':'Supplies','توريد':'Supplies',
          'خدمة':'Services','خدمات':'Services','ايجار':'Rental','إيجار':'Rental',
          'نقل':'Transport','مقاولة':'Subcontract','عقد':'Contract'};
  var ENG={'rental':'Rental','rentel':'Rental','rent':'Rental','supply':'Supplies','supplies':'Supplies',
           'supply & installation':'Supply & Installation','installation':'Installation','service':'Services',
           'services':'Services','labor':'Labor','labour':'Labor','transport':'Transport','transportation':'Transport',
           'workmanship':'Workmanship','subcontract':'Subcontract','contract':'Contract'};
  function ctype(v){if(v==null||typeof v==='number')return 'Unspecified';var s=String(v).trim();
    if(s==='')return 'Unspecified'; if(CT[s])return CT[s]; return ENG[s.toLowerCase()]||s;}
  var OST={'تم التوريد':'Delivered','تم النوريد':'Delivered','جاري':'In Progress','جارى':'In Progress',
           'تم ارتجاع بعض الخامات':'Partial Return','امر توريد':'Issued'};

  function sheetHidden(wb,name){            // skip hidden / very-hidden sheets
    var arr=(wb.Workbook&&wb.Workbook.Sheets)||[];
    for(var i=0;i<arr.length;i++) if(arr[i].name===name) return (arr[i].Hidden||0)>0;
    return false;
  }
  function sheetRows(wb,name){              // null if missing or hidden
    if(sheetHidden(wb,name)) return null;
    var ws=wb.Sheets[name]; if(!ws) return null;
    return {ws:ws, rows:XLSX.utils.sheet_to_json(ws,{header:1,raw:true,cellDates:true,defval:null})};
  }
  function rowHidden(ws,r){ var R=ws['!rows']; return !!(R && R[r] && R[r].hidden); }
  var CONTRACT_SHEETS=['code 8','demerdash hospital','جوهرة العاصمه','نادى الرمايه الدولى'];
  var MSR_SHEETS=['AL NOBARIAH BRIDGE','El Basra Bridge','AL hAMAM ','el nubaria Culverts '];
  var ORDERS_SHEET='اوامر توريد ';

  function extractContracts(wb){
    var out=[];
    CONTRACT_SHEETS.forEach(function(sn){
      var sr=sheetRows(wb,sn); if(!sr) return; var ws=sr.ws, rows=sr.rows;
      for(var r=2;r<rows.length;r++){             // 1-based row 3 => index 2
        if(rowHidden(ws,r)) continue;             // skip hidden rows
        var row=rows[r]||[]; var g=function(i){return row[i]==null?null:row[i];};
        var contractor=clean(g(14)), tow=clean(g(4)), cname=clean(g(11));
        if(clean(g(0))==='Average') continue;
        if(!(contractor||tow||cname)) continue;   // drop blank/total rows
        out.push({
          project:projn(clean(g(1))||sn), received:clean(g(2)), reqNo:clean(g(3)),
          typeOfWork:tow, contractType:clean(g(6)), ctype:ctype(g(6)),
          draftDate:clean(g(7)), approvalDraft:clean(g(9)), diffDays:num(g(10)),
          contractingName:cname, contractingDate:clean(g(12)), contractor:contractor,
          amount:num(g(16)), advancePct:num(g(17)), advanceAmt:num(g(18)),
          tenderAmount:num(g(19)), priceType:clean(g(20)), variance:num(g(21)), savingPct:num(g(22))
        });
      }
    });
    return out;
  }
  function extractLog(wb){
    var msr=[], orders=[];
    MSR_SHEETS.forEach(function(sn){
      var sr=sheetRows(wb,sn); if(!sr) return; var ws=sr.ws, rows=sr.rows;
      for(var r=1;r<rows.length;r++){             // 1-based row 2 => index 1
        if(rowHidden(ws,r)) continue;             // skip hidden rows
        var row=rows[r]||[]; var g=function(i){return row[i]==null?null:row[i];};
        var item=clean(g(6)), supplier=clean(g(13)), status=clean(g(11));
        if(!(item||supplier||status)) continue;
        msr.push({
          project:(clean(g(0))||sn).toString().trim(), po:clean(g(1)), createdDate:clean(g(2)),
          msrNo:clean(g(3)), msrDate:clean(g(4)), item:item, unit:clean(g(7)),
          qReq:num(g(8)), qConv:num(g(9)), qRequired:num(g(10)), status:status, supplier:supplier
        });
      }
    });
    var osr=sheetRows(wb,ORDERS_SHEET); var payPaid=0, payDue=0;
    if(osr){ var ows=osr.ws, orows=osr.rows;
      for(var r=2;r<orows.length;r++){            // 1-based row 3 => index 2
        if(rowHidden(ows,r)) continue;            // skip hidden rows
        var row=orows[r]||[]; var g=function(i){return row[i]==null?null:row[i];};
        payPaid+=num(g(10))||0; payDue+=num(g(11))||0;          // capture paid/due incl. any TOTAL row
        if(!(clean(g(1))||clean(g(4))||clean(g(5)))) continue;   // drop amount-only TOTAL rows
        var disc=num(g(8)), st=clean(g(6));
        // cols: 8=Amount 9=Discount 10=Discounted Amount 11=Paid 12=Due Amount
        orders.push({
          project:clean(g(0)), po:clean(g(1)), createdDate:clean(g(2)), msrDate:clean(g(3)),
          supplier:clean(g(4)), items:clean(g(5)), status:st,
          amount:num(g(7)), discount:disc, discountedAmount:num(g(9)), paid:num(g(10)), due:num(g(11)),
          statusEn:OST[(st||'').toString().trim()]||'Other', totalSaving:disc||0
        });
      }
    }
    return {orders:orders, msr:msr, payTotals:{paid:payPaid, due:payDue}};
  }

  function hasOrders(wb){
    if(wb.SheetNames.indexOf(ORDERS_SHEET)>=0) return true;
    for(var i=0;i<MSR_SHEETS.length;i++) if(wb.SheetNames.indexOf(MSR_SHEETS[i])>=0) return true;
    return false;
  }
  function classify(wb,name){                 // for the manual file-picker (no explicit type)
    name=(name||'').toLowerCase();
    if(hasOrders(wb)) return 'purchase';
    if(/equip|indirect/.test(name)) return 'equipment';
    return 'subcontract';
  }
  window.XLSXExtract={
    classify:classify,
    fromTyped:function(items){                 // items: [{wb, type?, name?}]
      var contracts=[], orders=[], msr=[], equipment=[], payTotals={paid:0,due:0};
      items.forEach(function(it){
        var t=it.type||classify(it.wb,it.name);
        if(t==='purchase'){ var l=extractLog(it.wb); orders=orders.concat(l.orders); msr=msr.concat(l.msr); payTotals.paid+=l.payTotals.paid; payTotals.due+=l.payTotals.due; }
        else if(t==='equipment'){ equipment=equipment.concat(extractContracts(it.wb)); }
        else { contracts=contracts.concat(extractContracts(it.wb)); }
      });
      return {contracts:contracts, orders:orders, msr:msr, equipment:equipment, payTotals:payTotals};
    },
    fromWorkbooks:function(wbs){ return this.fromTyped(wbs.map(function(w){return {wb:w};})); }
  };
})();
