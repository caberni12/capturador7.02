const API="https://script.google.com/macros/s/AKfycbzC_qrSyXeTw9NcO40ap4x2cfs3FZIBKqMZLV9kKhYYh7n2XTPAuj1Vb2ckpFBWi8Ys/exec";

let productos=[];
let capturas=JSON.parse(localStorage.getItem("capturas")||"[]");
let scanner=null,modo=null,torch=false,editIndex=-1;

operador.value=localStorage.getItem("operador")||"";
ubicacion.value=localStorage.getItem("ubicacion")||"";

fetch(API).then(r=>r.json()).then(d=>{
 productos=d;
 localStorage.setItem("productos",JSON.stringify(d));
}).catch(()=>{
 const c=localStorage.getItem("productos");
 if(c) productos=JSON.parse(c);
});

render();

function openTab(id){
 document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
 document.getElementById(id).classList.add('active');
}

function limpiarUbicacion(){
 ubicacion.value="";
 localStorage.removeItem("ubicacion");
 previewIngreso();
}

function buscarDescripcion(){
 const c=codigo.value.trim().toLowerCase();
 const p=productos.find(x=>String(x.CODIGO).toLowerCase()===c);
 if(p) descripcion.value=p.DESCRIPCION||"";
}

function previewIngreso(){
 if(!codigo.value && !descripcion.value){preview.innerHTML="";return;}
 preview.innerHTML=`<div class='row preview'><b>🕒 PREVISUALIZANDO</b><br><br>
 <b>${codigo.value||"-"}</b> – ${descripcion.value||"-"}<br>
 <span class='small'>${ubicacion.value||"SIN UBICACIÓN"} | ${operador.value||"-"} | Cant: ${cantidad.value}</span></div>`;
}

function scanCodigo(){modo="codigo";abrirScanner();}
function scanUbicacion(){modo="ubicacion";abrirScanner();}
function toggleScanner(){scannerBox.style.display==="none"?abrirScanner():cerrarScanner();}

function abrirScanner(){
 if(scanner) return;
 scannerBox.style.display="block";
 scanner=new Html5Qrcode("scannerBox");
 scanner.start(
  {facingMode:"environment"},
  {
    fps:12,
    qrbox:260,
    formatsToSupport:[
      Html5QrcodeSupportedFormats.QR_CODE,
      Html5QrcodeSupportedFormats.CODE_128,
      Html5QrcodeSupportedFormats.CODE_39,
      Html5QrcodeSupportedFormats.EAN_13
    ]
  },
  txt=>{
  beep.play();navigator.vibrate?.(200);
  if(modo==="codigo"){codigo.value=txt;buscarDescripcion();previewIngreso();}
  if(modo==="ubicacion"){ubicacion.value=txt;localStorage.setItem("ubicacion",txt);previewIngreso();}
  cerrarScanner();
 });
}
function cerrarScanner(){
 if(!scanner) return;
 scanner.stop().then(()=>{scanner.clear();scanner=null;scannerBox.style.display="none";});
}
function toggleTorch(){
 torch=!torch;
 scanner?.applyVideoConstraints({advanced:[{torch}]}).catch(()=>{});
}

function ingresar(){
 if(!codigo.value.trim()){
  alert("❌ Los datos no se pueden guardar. Digite un código correcto.");
  return;
 }
 localStorage.setItem("operador",operador.value);
 if(ubicacion.value) localStorage.setItem("ubicacion",ubicacion.value);
 else localStorage.removeItem("ubicacion");

 const d={
  Fecha:new Date().toLocaleString(),
  Operador:operador.value||"",
  Ubicación:ubicacion.value||"SIN UBICACIÓN",
  Código:codigo.value,
  Descripción:descripcion.value,
  Cantidad:Number(cantidad.value)
 };

 if(editIndex>=0){capturas[editIndex]=d;editIndex=-1;}
 else capturas.push(d);

 localStorage.setItem("capturas",JSON.stringify(capturas));
 limpiar();render();
}

function cargarParaEditar(i){
 const c=capturas[i];
 operador.value=c.Operador;
 ubicacion.value=c.Ubicación==="SIN UBICACIÓN"?"":c.Ubicación;
 codigo.value=c.Código;
 descripcion.value=c.Descripción;
 cantidad.value=c.Cantidad;
 editIndex=i;
 previewIngreso();
 render();
 window.scrollTo({top:0,behavior:"smooth"});
}

function cancelarEdicion(){
 editIndex=-1;
 limpiar();
 render();
}

function limpiar(){
 codigo.value="";descripcion.value="";cantidad.value=1;preview.innerHTML="";
}

function render(){
 tabla.innerHTML="";
 let total=0;
 capturas.forEach((c,i)=>{
  total+=Number(c.Cantidad)||0;
  tabla.innerHTML+=`<div class='row ${editIndex===i?"editing":""}'>
   <button class='delbtn' onclick='event.stopPropagation();eliminarItem(${i})'>×</button>
   <div onclick='cargarParaEditar(${i})'>
    <b>${c.Código}</b> – ${c.Descripción}<br>
    <span class='small'>${c.Ubicación} | ${c.Operador} | ${c.Fecha} | Cant: ${c.Cantidad}</span>
   </div>
  </div>`;
 });
 totalizador.innerText="Total unidades: "+total;
}

async function finalizar(){
 /* EXPORTA EXCEL Y DEJA LISTO PARA NUEVA CAPTURA */
 if(!capturas.length) return;
 const capturasExcel = capturas.map(r => ({
  ...r,
  Código: "'" + String(r.Código)
}));
const ws = XLSX.utils.json_to_sheet(capturasExcel);
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,ws,"Captura");
 const data=XLSX.write(wb,{bookType:"xlsx",type:"array"});

 if(window.showDirectoryPicker){
  try{
   const root=await window.showDirectoryPicker();
   const dir=await root.getDirectoryHandle("inventariosistema",{create:true});
   const file=await dir.getFileHandle("captura.xlsx",{create:true});
   const w=await file.createWritable();
   await w.write(data);await w.close();
  }catch(e){}
 }

 const blob=new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
 const url=URL.createObjectURL(blob);
 const a=document.createElement("a");
 a.href=url;a.download="captura.xlsx";a.target="_blank";a.click();

 /* LIMPIEZA TOTAL PARA NUEVA CAPTURA */
 localStorage.removeItem("capturas");
 capturas=[];
 limpiar();
 render();
 operador.value="";
 codigo.value="";
 descripcion.value="";
 cantidad.value=1;
 editIndex=-1;

}

function importarMaestra(){
 const file=fileExcel.files[0];
 if(!file) return alert("Selecciona Excel");
 const reader=new FileReader();
 reader.onload=e=>{
  const wb=XLSX.read(e.target.result,{type:"binary"});
  const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  enviarMaestra(data);
 };
 reader.readAsBinaryString(file);
}

async function enviarMaestra(data){
 barra.style.width="0%";
 mensaje.innerText="⏳ Importando...";
 let p=0;
 const t=setInterval(()=>{p+=10;barra.style.width=p+"%"},200);
 try{
  await fetch(API,{method:"POST",body:JSON.stringify({accion:"importar",data})});
  clearInterval(t);
  barra.style.width="100%";
  mensaje.innerText="✅ Importación exitosa";
  productos=data;
 }catch(e){
  clearInterval(t);
  mensaje.innerText="❌ Error al importar";
 }
}

function eliminarItem(i){
 if(!confirm("¿Eliminar este registro?")) return;
 capturas.splice(i,1);
 localStorage.setItem("capturas",JSON.stringify(capturas));
 if(editIndex===i) editIndex=-1;
 render();
}


function exportarPDF(){
 if(!capturas.length) return alert("Sin datos");
 const w = window.open("");
 let h = "<h3>Reporte de Captura</h3><table border='1' cellpadding='5' cellspacing='0'><tr>";
 Object.keys(capturas[0]).forEach(k=>h+="<th>"+k+"</th>");
 h+="</tr>";
 capturas.forEach(r=>{
  h+="<tr>";
  Object.values(r).forEach(v=>h+="<td>"+v+"</td>");
  h+="</tr>";
 });
 h+="</table>";
 w.document.write(h);
 w.print();
}

/* ===== MODAL CONSULTA SOLO VISUAL ===== */

let timerConsulta=null;
let filasConsulta=[];
let indexConsulta=-1;

function abrirModalConsulta(){
  modalConsulta.classList.add("show");
  buscarConsulta.value="";
  resultadoConsulta.innerHTML="";
  scrollConsulta.style.display="none";
  msgConsulta.innerText="Escriba para consultar";
  filasConsulta=[];
  indexConsulta=-1;
  buscarConsulta.focus();
}

function cerrarModalConsulta(){
  modalConsulta.classList.remove("show");
}

/* FILTRO DINÁMICO */
function filtrarConsulta(){
  clearTimeout(timerConsulta);
  timerConsulta=setTimeout(filtrarConsultaReal,300);
}

function filtrarConsultaReal(){
  const q=buscarConsulta.value.trim().toLowerCase();
  resultadoConsulta.innerHTML="";
  scrollConsulta.style.display="none";
  msgConsulta.innerText="";
  filasConsulta=[];
  indexConsulta=-1;

  if(q.length<2){
    msgConsulta.innerText="Escriba al menos 2 caracteres";
    return;
  }

  let count=0;
  for(const p of productos){   // usa tu maestra existente
    if(
      String(p.CODIGO).toLowerCase().includes(q) ||
      String(p.DESCRIPCION).toLowerCase().includes(q)
    ){
      const tr=document.createElement("tr");
      tr.innerHTML=`<td>${p.CODIGO}</td><td>${p.DESCRIPCION}</td>`;
      tr.onclick=()=>activarFilaConsulta(filasConsulta.length);
      resultadoConsulta.appendChild(tr);
      filasConsulta.push(tr);
      count++;
      if(count>=50) break;
    }
  }

  if(!count){
    msgConsulta.innerText="❌ Sin coincidencias";
    return;
  }

  scrollConsulta.style.display="block";
  activarFilaConsulta(0);
}

/* ACTIVAR FILA (CURSOR VISUAL + SCROLL) */
function activarFilaConsulta(i){
  if(i<0||i>=filasConsulta.length) return;

  filasConsulta.forEach(r=>r.classList.remove("selected"));
  filasConsulta[i].classList.add("selected");
  indexConsulta=i;

  const fila=filasConsulta[i];
  const cont=scrollConsulta;

  const filaTop=fila.offsetTop;
  const filaBottom=filaTop+fila.offsetHeight;
  const contTop=cont.scrollTop;
  const contBottom=contTop+cont.clientHeight;

  if(filaTop<contTop){
    cont.scrollTop=filaTop-10;
  }else if(filaBottom>contBottom){
    cont.scrollTop=filaBottom-cont.clientHeight+10;
  }
}

/* ===== TECLADO ===== */
document.addEventListener("keydown",e=>{
  if(!modalConsulta.classList.contains("show")) return;
  if(!filasConsulta.length) return;

  if(e.key==="ArrowDown"){
    e.preventDefault();
    activarFilaConsulta(Math.min(indexConsulta+1,filasConsulta.length-1));
  }

  if(e.key==="ArrowUp"){
    e.preventDefault();
    activarFilaConsulta(Math.max(indexConsulta-1,0));
  }

  if(e.key==="Escape"){
    cerrarModalConsulta();
  }
});

/* ===== EXPORTAR MAESTRA DE PRODUCTOS ===== */
function exportarMaestraProductos(){

if(!productos || !productos.length){
  alert("❌ No hay productos para exportar");
  return;
}

// Fuerza CODIGO como texto (muy importante)
const data = productos.map(p => ({
  CODIGO: "" + String(p.CODIGO),
  DESCRIPCION: p.DESCRIPCION
}));

const ws = XLSX.utils.json_to_sheet(data);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Maestra_Productos");

const excel = XLSX.write(wb, {
  bookType: "xlsx",
  type: "array"
});

const blob = new Blob([excel], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});

const url = URL.createObjectURL(blob);
const a = document.createElement("a");
a.href = url;
a.download = "maestra_productos.xlsx";
a.click();

URL.revokeObjectURL(url);
}

