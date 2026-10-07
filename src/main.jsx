import React, {useMemo, useRef, useState} from 'react';
import {createRoot} from 'react-dom/client';
import {BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, Treemap, LineChart, Line} from 'recharts';
import * as XLSX from 'xlsx';
import {Upload, Download, Mail, RotateCcw, Search, ChevronLeft, ChevronRight, FileSpreadsheet, TrendingUp, WalletCards, Layers3, AlertTriangle, CheckCircle2, X, MousePointer2} from 'lucide-react';
import './styles.css';

const MONTHS=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const COLORS=['#163A5F','#2F6690','#D4A017','#5B7C99','#8C6A2B','#3D5A73','#B08D2C','#244B68'];
const BRAND={navy:'#0F2438',blue:'#163A5F',cyan:'#2F6690',lime:'#D4A017',amber:'#B88912',bg:'#F3F5F7',card:'#FFFFFF',text:'#17212B',muted:'#687582',border:'#DDE3E8'};

function num(v){ if(v===null||v===undefined||v==='') return 0; if(typeof v==='number') return Number.isFinite(v)?v:0; const s=String(v).replace(/\./g,'').replace(',','.').replace(/[^0-9.-]/g,''); return Number(s)||0; }
function money(v){ return new Intl.NumberFormat('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2}).format(v||0); }
function moneyFull(v,currency='Bs.'){ return `${currency} ${money(v)}`; }
function compact(v){ return new Intl.NumberFormat('es-ES',{maximumFractionDigits:0}).format(v||0); }
function firstHeader(rows){ return rows.findIndex(r=>Array.isArray(r)&&r.some(x=>String(x??'').trim().toLowerCase()==='cuenta')); }
function parseWorkbook(buffer){
  const wb=XLSX.read(buffer,{type:'array',cellDates:true});
  const sheet=wb.Sheets[wb.SheetNames[0]]; const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null});
  const h=firstHeader(rows); if(h<0) throw new Error('No se encontró una fila con la columna Cuenta.');
  const headers=rows[h].map(x=>String(x??'').trim());
  const data=rows.slice(h+1).filter(r=>r[0]!==null&&String(r[0]).trim()!==''&&String(r[1]??'').trim().toLowerCase()!=='totales:').map(r=>{
    const o={}; headers.forEach((key,i)=>o[key||`Columna ${i+1}`]=r[i]);
    const monthsObj=Object.fromEntries(MONTHS.map((m,i)=>[m,num(r[i+3])]));
    const sumMonths=MONTHS.reduce((s,m,i)=>s+num(r[i+3]),0);
    const totalVal=(r[15]!==undefined&&r[15]!==null&&String(r[15]).trim()!=='')?num(r[15]):sumMonths;
    return {Cuenta:String(r[0]),Descripcion:String(r[1]??''),SaldoInicial:num(r[2]),...monthsObj,Total:totalVal,Nivel:String(r[0]).split('.').length,Tipo:String(r[0]).includes('.')?'Detalle':'Grupo'};
  });
  return {rows:data,meta:{sheet:wb.SheetNames[0],periodo:rows[1]?.[0]||'',empresa:rows[3]?.[0]||'',rif:rows[2]?.[0]||'',sheets:wb.SheetNames}};
}
function downloadBlob(name,blob){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),500);}
function exportWorkbook(rows,meta){
  const out=rows.map(({Nivel,Tipo,...r})=>r); const ws=XLSX.utils.json_to_sheet(out); const summary=[['Dashboard Financiero'],['Empresa',meta.empresa],['Periodo',meta.periodo],['Registros',rows.length],['Total seleccionado',rows.reduce((s,r)=>s+r.Total,0)]]; const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),'Resumen'); XLSX.utils.book_append_sheet(wb,ws,'Detalle'); const b=XLSX.write(wb,{bookType:'xlsx',type:'array'}); downloadBlob('Financiero_Dashboard_Export.xlsx',new Blob([b],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); }

function downloadTemplate(){
  const headerMeta = [
    ['Saldos de Cuentas Mensualizado'],
    ['Período: Enero a Diciembre 2026'],
    ['RIF: J-00000000-0'],
    ['EMPRESA EJEMPLO, C.A.'],
    ['']
  ];
  const colHeaders = ['Cuenta', 'Descripción', 'Saldo Inicial', ...MONTHS, 'Total'];
  const sampleRows = [
    ['1', 'ACTIVO', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['1.01', 'ACTIVO CORRIENTE', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['1.01.01', 'Caja y Bancos', 15000, 2500, 3100, 2800, 3200, 2900, 3400, 3800, 3600, 4000, 4200, 4500, 4800, 42800],
    ['1.01.02', 'Cuentas por Cobrar Comerciales', 12000, 1800, 2200, 2100, 2400, 2300, 2600, 2800, 2700, 3000, 3200, 3500, 3800, 32400],
    ['1.02', 'ACTIVO NO CORRIENTE', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['1.02.01', 'Mobiliario y Equipos', 8000, 0, 0, 1200, 0, 0, 0, 1500, 0, 0, 0, 2000, 0, 4700],
    ['2', 'PASIVO', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['2.01', 'PASIVO CORRIENTE', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['2.01.01', 'Cuentas por Pagar Proveedores', 8500, 1400, 1600, 1550, 1700, 1650, 1800, 1900, 1850, 2000, 2100, 2200, 2400, 22150],
    ['2.01.02', 'Impuestos y Retenciones por Pagar', 2100, 450, 520, 480, 560, 510, 590, 620, 600, 650, 680, 710, 750, 7120],
    ['3', 'PATRIMONIO', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['3.01.01', 'Capital Social Suscrito y Pagado', 25000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['3.02.01', 'Resultados Acumulados', 5000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['4', 'INGRESOS', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['4.01.01', 'Ingresos por Ventas de Bienes', 0, 12500, 14200, 13800, 15600, 15100, 16800, 17500, 17100, 18400, 19200, 20500, 22000, 202700],
    ['4.01.02', 'Ingresos por Servicios Prestados', 0, 3200, 3600, 3400, 3900, 3700, 4100, 4300, 4200, 4500, 4700, 5000, 5300, 49900],
    ['5', 'COSTOS', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['5.01.01', 'Costo de Ventas y Mercancías', 0, 6200, 7100, 6900, 7800, 7550, 8400, 8750, 8550, 9200, 9600, 10250, 11000, 101300],
    ['6', 'GASTOS', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ['6.01.01', 'Gastos de Personal (Sueldos)', 0, 3500, 3500, 3500, 3800, 3800, 3800, 4100, 4100, 4100, 4400, 4400, 5500, 48000],
    ['6.01.02', 'Servicios Públicos y Generales', 0, 650, 720, 680, 750, 710, 780, 820, 790, 850, 880, 920, 980, 9530],
    ['6.01.03', 'Alquileres y Mantenimiento', 0, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 1200, 14400]
  ];
  const fullAoa = [...headerMeta, colHeaders, ...sampleRows];
  const ws = XLSX.utils.aoa_to_sheet(fullAoa);
  ws['!cols'] = [
    { wch: 14 },
    { wch: 38 },
    { wch: 15 },
    ...MONTHS.map(() => ({ wch: 13 })),
    { wch: 16 }
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'SaldosMensualizados');
  const b = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  downloadBlob('Plantilla_Saldos_Mensualizados.xlsx', new Blob([b], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
}

const STANDARD_GROUPS = {
  '1': 'Activos',
  '2': 'Pasivos',
  '3': 'Patrimonio',
  '4': 'Ingresos',
  '5': 'Costos',
  '6': 'Gastos',
  '7': 'Otros Ingresos y Egresos',
  '8': 'Cuentas de Orden',
  '9': 'Cuentas Fiscales y de Control'
};

function getAccountGroupCode(cuenta) {
  const s = String(cuenta || '').trim();
  if (!s) return '1';
  const sepMatch = s.match(/^(\d+)[.\-_]/);
  if (sepMatch) return sepMatch[1];
  const numMatch = s.match(/^\d+/);
  if (numMatch) return numMatch[0].charAt(0);
  return s.split(/[.\-_ ]/)[0] || '1';
}

function getGroupName(groupCode, allRows = []) {
  const code = String(groupCode || '').trim();
  if (!code) return 'Sin Grupo';
  const rootRow = allRows.find(r => {
    const c = String(r.Cuenta || '').trim();
    return c === code || c === `${code}.` || (r.Nivel === 1 && c.startsWith(code));
  });
  if (rootRow && rootRow.Descripcion && rootRow.Descripcion.trim()) {
    const rawDesc = rootRow.Descripcion.trim();
    const isAllUpper = rawDesc === rawDesc.toUpperCase() && rawDesc.length > 2;
    const formattedDesc = isAllUpper 
      ? rawDesc.charAt(0).toUpperCase() + rawDesc.slice(1).toLowerCase() 
      : rawDesc;
    return `${code} - ${formattedDesc}`;
  }
  if (STANDARD_GROUPS[code]) {
    return `${code} - ${STANDARD_GROUPS[code]}`;
  }
  return `${code} - Grupo ${code}`;
}

function App(){
 const inputRef=useRef(); const [data,setData]=useState(null); const [tab,setTab]=useState('resumen'); const [filters,setFilters]=useState({nivel:'Todos',cuenta:'',minTotal:'',maxTotal:'',mes:'Todos'}); const [selected,setSelected]=useState(null); const [sort,setSort]=useState({key:'Total',dir:'desc'}); const [page,setPage]=useState(1); const pageSize=16;
 const load=async file=>{if(!file)return; try{const buf=await file.arrayBuffer();setData(parseWorkbook(buf));setSelected(null);setFilters({nivel:'Todos',cuenta:'',minTotal:'',maxTotal:'',mes:'Todos'});setPage(1);}catch(e){alert(e.message)}};
 const filtered=useMemo(()=>{if(!data)return[];return data.rows.filter(r=>(filters.nivel==='Todos'||String(r.Nivel)===filters.nivel)&&(filters.mes==='Todos'||r[filters.mes]!==0)&&(!filters.cuenta||r.Cuenta.toLowerCase().includes(filters.cuenta.toLowerCase())||r.Descripcion.toLowerCase().includes(filters.cuenta.toLowerCase()))&&(filters.minTotal===''||r.Total>=num(filters.minTotal))&&(filters.maxTotal===''||r.Total<=num(filters.maxTotal)));},[data,filters]);
 
 // Monthly evolution: sum accounts by accounting group (1 - Activos, 2 - Pasivos, ...) sorted 1, 2, 3...
 const monthlyByGroup=useMemo(()=>{
   if(!filtered.length) return {groups:[], data:[]};
   const groupCodesSet=new Set();
   filtered.forEach(r=>{
     const code=getAccountGroupCode(r.Cuenta);
     if(code) groupCodesSet.add(code);
   });
   const sortedCodes=[...groupCodesSet].sort((a,b)=>(parseInt(a,10)||0)-(parseInt(b,10)||0));
   const groupDefs=sortedCodes.map(code=>({
     code,
     name:getGroupName(code, data?.rows)
   }));
   const chartData=MONTHS.map(m=>{
     const pt={mes:m.slice(0,3)};
     groupDefs.forEach(g=>{
       const sumMonth=filtered
         .filter(r=>getAccountGroupCode(r.Cuenta)===g.code)
         .reduce((sum,r)=>sum+(r[m]||0),0);
       pt[g.name]=sumMonth;
     });
     return pt;
   });
   return {groups:groupDefs, data:chartData};
 },[filtered, data]);

 const monthly=useMemo(()=>MONTHS.map(m=>({mes:m.slice(0,3),valor:filtered.reduce((s,r)=>s+r[m],0)})),[filtered]);
 
 // Group distribution: sum by group (1 - Activos, 2 - Pasivos, ...) sorted 1, 2, 3...
 const groups=useMemo(()=>{
   if(!filtered.length) return [];
   const map=new Map();
   filtered.forEach(r=>{
     const code=getAccountGroupCode(r.Cuenta);
     if(!code) return;
     map.set(code,(map.get(code)||0)+(r.Total||0));
   });
   return [...map.entries()]
     .map(([code,value])=>({
       code,
       name:getGroupName(code, data?.rows),
       value,
       pieValue: Math.max(Math.abs(value), 0.0001)
     }))
     .sort((a,b)=>(parseInt(a.code,10)||0)-(parseInt(b.code,10)||0));
 },[filtered, data]);

 const top=useMemo(()=>[...filtered].sort((a,b)=>Math.abs(b.Total)-Math.abs(a.Total)).slice(0,10),[filtered]);
 const tree=useMemo(()=>top.map((r,i)=>({name:`${r.Cuenta} ${r.Descripcion}`,size:Math.max(Math.abs(r.Total),1),value:r.Total,Cuenta:r.Cuenta,index:i})),[top]);
 const total=filtered.reduce((s,r)=>s+r.Total,0), initial=filtered.reduce((s,r)=>s+r.SaldoInicial,0), abs=filtered.reduce((s,r)=>s+Math.abs(r.Total),0), positive=filtered.filter(r=>r.Total>0).length, negative=filtered.filter(r=>r.Total<0).length;
 // Sort detail by group (first segment of account code), then by full account code
 const sorted=useMemo(()=>{
   const base=[...filtered];
   if(sort.key==='Cuenta'||sort.key==='Descripcion'||sort.key==='SaldoInicial'||sort.key==='Total'||MONTHS.includes(sort.key)){
     // primary: group (first segment), secondary: user sort
     base.sort((a,b)=>{
       const ga=getAccountGroupCode(a.Cuenta), gb=getAccountGroupCode(b.Cuenta);
       if(ga!==gb) return ga.localeCompare(gb,undefined,{numeric:true});
       const av=a[sort.key], bv=b[sort.key];
       return (av>bv?1:av<bv?-1:0)*(sort.dir==='asc'?1:-1);
     });
   } else {
     base.sort((a,b)=>{
       const ga=getAccountGroupCode(a.Cuenta), gb=getAccountGroupCode(b.Cuenta);
       if(ga!==gb) return ga.localeCompare(gb,undefined,{numeric:true});
       return a.Cuenta.localeCompare(b.Cuenta,undefined,{numeric:true});
     });
   }
   return base;
 },[filtered,sort]);
 const pages=Math.max(1,Math.ceil(sorted.length/pageSize)); const pageRows=sorted.slice((page-1)*pageSize,page*pageSize);
 const clickFilter=(kind,value)=>{if(kind==='cuenta')setFilters(f=>({...f,cuenta:value}));else if(kind==='mes')setFilters(f=>({...f,mes:value}));else if(kind==='nivel')setFilters(f=>({...f,nivel:value}));setTab('detalle');setPage(1)};
 const reset=()=>{setFilters({nivel:'Todos',cuenta:'',minTotal:'',maxTotal:'',mes:'Todos'});setSelected(null);setPage(1)};
 const active=[filters.nivel!=='Todos'?['Nivel',filters.nivel]:null,filters.cuenta?['Cuenta',filters.cuenta]:null,filters.minTotal!==''?['Mín.',money(filters.minTotal)]:null,filters.maxTotal!==''?['Máx.',money(filters.maxTotal)]:null,filters.mes!=='Todos'?['Mes',filters.mes]:null].filter(Boolean);
 const conclusions=useMemo(()=>{if(!filtered.length)return['No hay registros con los filtros actuales.'];const max=top[0];const peak=monthly.reduce((a,b)=>Math.abs(b.valor)>Math.abs(a.valor)?b:a,monthly[0]);return [`La cartera filtrada reúne ${compact(filtered.length)} cuentas y un total neto de ${moneyFull(total)}.`,`La mayor concentración por importe absoluto está en ${max?.Cuenta||'—'} · ${max?.Descripcion||'—'}, con ${moneyFull(max?.Total||0)}.`,`El movimiento mensual de mayor magnitud es ${peak?.mes||'—'}, con ${moneyFull(peak?.valor||0)}; conviene revisar los asientos de ese periodo.`,`Hay ${compact(positive)} cuentas con saldo total positivo y ${compact(negative)} con saldo negativo; la mezcla permite priorizar conciliación y revisión.`,`El treemap muestra la concentración relativa de las cuentas y los clics sobre las visualizaciones aplican filtros al detalle.`]},[filtered,total,top,monthly,positive,negative]);
 const onDrop=e=>{e.preventDefault();load(e.dataTransfer.files?.[0])};
 return <div className="app" onDragOver={e=>e.preventDefault()} onDrop={onDrop}>
   <aside className="sidebar">
    <div className="brand"><div className="brandMark">F</div><div><b>FINANCIERODATA</b><span>Análisis · Control · Decisiones</span></div></div>
    <div className="actions">
      <button className="action primary" onClick={()=>inputRef.current?.click()}><Upload size={15}/> Cargar Excel</button>
      <button className="action" onClick={downloadTemplate} title="Descargar plantilla Excel para rellenar"><FileSpreadsheet size={15}/> Plantilla Excel</button>
      <button className="action" disabled={!data} onClick={()=>exportWorkbook(filtered,data?.meta)}><Download size={15}/> Exportar Excel</button>
      <button className="action" disabled={!data} onClick={()=>{const body=encodeURIComponent(`Resumen Financiero\n${data.meta.empresa}\n${data.meta.periodo}\nRegistros: ${filtered.length}\nTotal: ${moneyFull(total)}\n\nConclusiones:\n- ${conclusions.join('\n- ')}`);location.href=`mailto:?subject=${encodeURIComponent('Borrador · Dashboard Financiero')}&body=${body}`}}><Mail size={15}/> Borrador por Email</button>
      <input ref={inputRef} hidden type="file" accept=".xlsx,.xls,.csv" onChange={e=>load(e.target.files?.[0])}/>
    </div>
    <nav><button className={tab==='resumen'?'active':''} onClick={()=>setTab('resumen')}>▦ <span>Resumen General</span></button><button className={tab==='categorias'?'active':''} onClick={()=>setTab('categorias')}>◫ <span>Análisis de Categorías</span></button><button className={tab==='detalle'?'active':''} onClick={()=>setTab('detalle')}>☷ <span>Detalle de Registros</span></button></nav>
    <div className="sidebarNote"><MousePointer2 size={14}/><span>Haz clic en cualquier gráfico para filtrar el detalle.</span></div>
    <div className="privacy"><CheckCircle2 size={13}/> Procesamiento 100% local</div>
   </aside>
   <main className="main">
    <header className="top"><div><div className="eyebrow">DASHBOARD CONTABLE</div><h1>{data?.meta.empresa||'Análisis de Saldos de Cuentas'}</h1><p>{data?.meta.periodo||'Carga un Excel para comenzar el análisis'}</p></div><div className="filters"><div className="filterRow"><select value={filters.nivel} onChange={e=>{setFilters(f=>({...f,nivel:e.target.value}));setPage(1)}}><option>Todos</option>{[...new Set(data?.rows.map(r=>r.Nivel)||[])].sort().map(n=><option key={n}>{n}</option>)}</select><div className="search"><Search size={14}/><input placeholder="Buscar cuenta o descripción…" value={filters.cuenta} onChange={e=>{setFilters(f=>({...f,cuenta:e.target.value}));setPage(1)}}/></div><input className="mini" type="number" placeholder="Mín. total" value={filters.minTotal} onChange={e=>setFilters(f=>({...f,minTotal:e.target.value}))}/><input className="mini" type="number" placeholder="Máx. total" value={filters.maxTotal} onChange={e=>setFilters(f=>({...f,maxTotal:e.target.value}))}/></div><div className="chips">{active.map(([k,v])=><span className="chip" key={k}>{k}: {v}<X size={12} onClick={()=>{if(k==='Nivel')setFilters(f=>({...f,nivel:'Todos'}));if(k==='Cuenta')setFilters(f=>({...f,cuenta:''}));if(k==='Mín.')setFilters(f=>({...f,minTotal:''}));if(k==='Máx.')setFilters(f=>({...f,maxTotal:''}));if(k==='Mes')setFilters(f=>({...f,mes:'Todos'}))}}/></span>)}{active.length>0&&<button className="reset" onClick={reset}><RotateCcw size={12}/> Restablecer</button>}</div></div></header>
    {!data?<Empty onLoad={()=>inputRef.current?.click()} onDownloadTemplate={downloadTemplate}/>:<>
      <section className="kpis"><Kpi icon={<WalletCards/>} label="Total neto" value={moneyFull(total)} tone="blue" sub="suma de Total filtrado"/><Kpi icon={<Layers3/>} label="Cuentas" value={compact(filtered.length)} tone="cyan" sub={`${compact(data.rows.length)} en archivo`}/><Kpi icon={<TrendingUp/>} label="Saldo inicial" value={moneyFull(initial)} tone="lime" sub="saldo inicial agregado"/><Kpi icon={<AlertTriangle/>} label="Concentración" value={moneyFull(abs)} tone="amber" sub="valor absoluto del total"/></section>
      {tab==='resumen'&&<div className="grid">
        <Card title="Evolución mensual" subtitle="Barras apiladas por grupo contable · Eje X: Meses · Eje Y: Monto · Grupos 1, 2, 3..." className="wide">
          <div style={{overflowX:'auto'}}>
            <div style={{minWidth:600}}>
              <ResponsiveContainer width="100%" height={280}>
                <BarChart data={monthlyByGroup.data} margin={{left:8,right:16,top:8,bottom:4}}
                  onClick={e=>e?.activeLabel&&clickFilter('mes',MONTHS.find(m=>m.slice(0,3)===e.activeLabel)||e.activeLabel)}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#EAECF0"/>
                  <XAxis dataKey="mes" tick={{fontSize:10}}/>
                  <YAxis tick={{fontSize:10}} tickFormatter={v=>money(v)} width={90}/>
                  <Tooltip
                    formatter={(v,name)=>[moneyFull(v), name]}
                    labelFormatter={l=>`Mes: ${l}`}
                    contentStyle={{fontSize:11,maxHeight:240,overflowY:'auto'}}
                  />
                  <Legend wrapperStyle={{fontSize:10,lineHeight:'18px'}}/>
                  {monthlyByGroup.groups.map((grp,i)=>(
                    <Bar
                      key={grp.code}
                      dataKey={grp.name}
                      name={grp.name}
                      stackId="monthly"
                      fill={COLORS[i%COLORS.length]}
                      onClick={()=>clickFilter('cuenta',grp.code)}
                      style={{cursor:'pointer'}}
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </Card>
        <Card title="Distribución por grupo" subtitle="Basado en columna Total final por grupo (1, 2, 3...)"><ResponsiveContainer width="100%" height={220}><PieChart><Pie data={groups} dataKey="pieValue" nameKey="name" innerRadius={55} outerRadius={85} paddingAngle={2} onClick={d=>clickFilter('cuenta',d.code||d.name.split(' - ')[0])}>{groups.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip formatter={(v,name,item)=>[moneyFull(item?.payload?.value??v),name]}/><Legend wrapperStyle={{fontSize:10}}/></PieChart></ResponsiveContainer></Card>
        <Card title="Top cuentas por importe absoluto" subtitle="Basado en columna Total final" className="wide"><ResponsiveContainer width="100%" height={220}><BarChart data={top.slice(0,8)} layout="vertical" margin={{left:12,right:16}}><CartesianGrid strokeDasharray="3 3" stroke="#EAECF0"/><XAxis type="number" tick={{fontSize:10}} tickFormatter={v=>money(v)}/><YAxis type="category" dataKey="Cuenta" width={82} tick={{fontSize:10}}/><Tooltip formatter={v=>moneyFull(v)}/><Bar dataKey="Total" radius={[0,4,4,0]} onClick={d=>clickFilter('cuenta',d.Cuenta)}>{top.slice(0,8).map((_,i)=><Cell key={i} fill={i===0?BRAND.amber:BRAND.blue}/>)}</Bar></BarChart></ResponsiveContainer></Card>
        <Card title="Mapa de cuentas" subtitle="Treemap · Columna Total final (muestra valores negativos)"><ResponsiveContainer width="100%" height={220}><Treemap data={tree} dataKey="size" stroke="#fff" content={<TreeContent onClick={r=>clickFilter('cuenta',r?.Cuenta||'')}/>} /></ResponsiveContainer></Card>
        <Card title="Conclusiones ejecutivas" subtitle="Lectura automática del estado filtrado" className="wide"><div className="conclusions">{conclusions.map((c,i)=><div key={i}><span>{String(i+1).padStart(2,'0')}</span><p>{c}</p></div>)}</div></Card>
      </div>}
      {tab==='categorias'&&<div className="grid categoryGrid"><Card title="Jerarquía de cuentas" subtitle="Concentración por código" className="wide"><ResponsiveContainer width="100%" height={300}><Treemap data={tree} dataKey="size" stroke="#fff" content={<TreeContent/>}/></ResponsiveContainer></Card><Card title="Grupos principales"><ResponsiveContainer width="100%" height={300}><BarChart data={groups} layout="vertical" margin={{left:10,right:20}}><CartesianGrid strokeDasharray="3 3" stroke="#EAECF0"/><XAxis type="number" tick={{fontSize:10}} tickFormatter={v=>money(v)}/><YAxis type="category" dataKey="name" width={120} tick={{fontSize:9}}/><Tooltip formatter={v=>moneyFull(v)}/><Bar dataKey="value" fill={BRAND.cyan} radius={[0,4,4,0]}/></BarChart></ResponsiveContainer></Card></div>}
      {tab==='detalle'&&<Card title="Detalle de registros" subtitle={`${filtered.length} registros · agrupados por grupo contable`} className="tableCard"><div className="tableWrap"><table><thead><tr>{['Cuenta','Descripcion','SaldoInicial',...MONTHS,'Total'].map(k=><th key={k} onClick={()=>{const dir=sort.key===k&&sort.dir==='desc'?'asc':'desc';setSort({key:k,dir})}}>{k==='SaldoInicial'?'Saldo inicial':k}{sort.key===k?' '+(sort.dir==='desc'?'↓':'↑'):''}</th>)}</tr></thead><tbody>{pageRows.map((r,idx)=>{
        const prevR=pageRows[idx-1];
        const curGroup=getAccountGroupCode(r.Cuenta);
        const prevGroup=prevR?getAccountGroupCode(prevR.Cuenta):null;
        const isGroupStart=curGroup!==prevGroup;
        return (
          <React.Fragment key={r.Cuenta}>
            {isGroupStart&&<tr className="groupHeader"><td colSpan={16+MONTHS.length}>
              <span className="groupLabel">{getGroupName(curGroup, data?.rows)}</span>
            </td></tr>}
            <tr className={selected===r.Cuenta?'selected':''} onClick={()=>setSelected(r.Cuenta)}>
              <td><b>{r.Cuenta}</b></td>
              <td title={r.Descripcion}>{r.Descripcion}</td>
              <td>{money(r.SaldoInicial)}</td>
              {MONTHS.map(m=><td key={m}>{money(r[m])}</td>)}
              <td className={r.Total<0?'negative':'positive'}>{money(r.Total)}</td>
            </tr>
          </React.Fragment>
        );
      })}</tbody><tfoot><tr className="totalsRow"><td><b>Totales:</b></td><td></td><td><b>{money(initial)}</b></td>{MONTHS.map(m=><td key={m}><b>{money(filtered.reduce((s,r)=>s+r[m],0))}</b></td>)}<td className={total<0?'negative':'positive'}><b>{money(total)}</b></td></tr></tfoot></table></div><div className="pager"><span>Página {page} de {pages}</span><button disabled={page<=1} onClick={()=>setPage(p=>p-1)}><ChevronLeft size={14}/></button><button disabled={page>=pages} onClick={()=>setPage(p=>p+1)}><ChevronRight size={14}/></button></div></Card>}
    </>}
   </main>
 </div>
}
function Kpi({icon,label,value,tone,sub}){return <div className={`kpi ${tone}`}><div className="kpiIcon">{icon}</div><div><span>{label}</span><strong>{value}</strong><small>{sub}</small></div></div>}
function Card({title,subtitle,children,className=''}){return <section className={`card ${className}`}><div className="cardHead"><div><h2>{title}</h2><span>{subtitle}</span></div></div>{children}</section>}
function TreeContent({x,y,width,height,name,value,index,Cuenta,onClick}){
  if(width<38||height<28)return null;
  const rawVal = typeof value === 'object' ? (value?.value ?? value?.size) : value;
  const val = rawVal ?? 0;
  const isNeg = val < 0;
  const bg = isNeg ? '#A33B36' : COLORS[(index||0)%COLORS.length];
  return <g onClick={()=>onClick?.(Cuenta||String(name).split(' ')[0])} style={{cursor:onClick?'pointer':'default'}}>
    <rect x={x} y={y} width={width} height={height} rx={5} fill={bg} opacity=".9" stroke="#fff"/>
    <text x={x+7} y={y+16} fill="#fff" fontSize="10" fontWeight="700">{String(name).slice(0,22)}</text>
    <text x={x+7} y={y+29} fill="#fff" fontSize="9" fontWeight="600">{money(val)}</text>
  </g>
}
function Empty({onLoad, onDownloadTemplate}){
  return <div className="empty">
    <div className="emptyIcon"><FileSpreadsheet size={32}/></div>
    <h2>Carga tu Excel para comenzar</h2>
    <p>Arrastra el archivo aquí o utiliza el botón de carga. Si necesitas el formato estándar, descarga la plantilla prediseñada lista para rellenar.</p>
    <div style={{display:'flex',gap:'10px',justifyContent:'center',flexWrap:'wrap',marginTop:'10px'}}>
      <button className="action primary big" onClick={onLoad}><Upload size={16}/> Cargar Excel</button>
      <button className="action big" onClick={onDownloadTemplate}><Download size={16}/> Descargar Plantilla Excel</button>
    </div>
    <div className="hint">Compatible con .xlsx, .xls y .csv · Incluye cuentas de muestra agrupadas del 1 al 6</div>
  </div>
}

createRoot(document.getElementById('root')).render(<App/>);
