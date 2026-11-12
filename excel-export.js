/* Exportação XLSX nativa, sem bibliotecas externas.
   Campos de acompanhamento são editáveis SOMENTE no Excel gerado. */
(function(){
'use strict';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const x=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
function crc32(data){let c=-1;for(const b of data){c^=b;for(let i=0;i<8;i++)c=(c>>>1)^((c&1)?0xedb88320:0);}return(c^-1)>>>0}
const utf=s=>new TextEncoder().encode(s);
function zip(entries){
 const local=[],central=[];let off=0;
 const w16=(view,pos,v)=>view.setUint16(pos,v,true),w32=(view,pos,v)=>view.setUint32(pos,v>>>0,true);
 for(const [path,xml] of entries){
  const name=utf(path),data=utf(xml),crc=crc32(data);
  const h=new Uint8Array(30+name.length),dv=new DataView(h.buffer);
  w32(dv,0,0x04034b50);w16(dv,4,20);w16(dv,6,0x0800);w16(dv,8,0);
  w32(dv,14,crc);w32(dv,18,data.length);w32(dv,22,data.length);w16(dv,26,name.length);
  h.set(name,30);local.push(h,data);
  const ch=new Uint8Array(46+name.length),cv=new DataView(ch.buffer);
  w32(cv,0,0x02014b50);w16(cv,4,20);w16(cv,6,20);w16(cv,8,0x0800);
  w32(cv,16,crc);w32(cv,20,data.length);w32(cv,24,data.length);w16(cv,28,name.length);w32(cv,42,off);
  ch.set(name,46);central.push(ch);off+=h.length+data.length;
 }
 const cdSize=central.reduce((n,a)=>n+a.length,0),e=new Uint8Array(22),ev=new DataView(e.buffer);
 w32(ev,0,0x06054b50);w16(ev,8,entries.length);w16(ev,10,entries.length);w32(ev,12,cdSize);w32(ev,16,off);
 return new Blob([...local,...central,e],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
const strCell=(ref,v,style)=>'<c r="'+ref+'" s="'+style+'" t="inlineStr"><is><t>'+x(v)+'</t></is></c>';
const numberCell=(ref,v,style)=>'<c r="'+ref+'" s="'+style+'"><v>'+Number(v)+'</v></c>';
const formulaCell=(ref,f,style,text=false)=>'<c r="'+ref+'" s="'+style+'"'+(text?' t="str"':'')+'><f>'+x(f)+'</f><v></v></c>';
function styles(){
 const font=(size,color,bold=false)=>'<font><sz val="'+size+'"/><name val="Aptos"/>'+(bold?'<b/>':'')+'<color rgb="'+color+'"/></font>';
 const fill=c=>'<fill><patternFill patternType="solid"><fgColor rgb="'+c+'"/><bgColor indexed="64"/></patternFill></fill>';
 const alignment=(kind='center')=>'<alignment horizontal="'+kind+'" vertical="center" wrapText="1"/>';
 const xf=(fontId,fillId,borderId=0,kind='center')=>'<xf numFmtId="0" fontId="'+fontId+'" fillId="'+fillId+'" borderId="'+borderId+'" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'+alignment(kind)+'</xf>';
 const colors={navy:'FF1D3447',teal:'FF0A6A6B',pale:'FFE6F3F1',ink:'FF2B3E4C',white:'FFFFFFFF',stripe:'FFF7FAFC',gold:'FFFFF0C9',blue:'FFEAF2F8',note:'FFFFF5DF',available:'FFE3F4E9',availableInk:'FF246641'};
 const fonts=[font(11,colors.ink),font(19,colors.white,true),font(10,colors.white),font(9,colors.teal,true),font(17,colors.navy,true),font(10,colors.ink),font(10,colors.white,true),font(10,colors.ink),font(10,colors.teal,true),font(10,colors.availableInk,true)].join('');
 const fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>',fill(colors.navy),fill(colors.teal),fill(colors.pale),fill(colors.gold),fill(colors.white),fill(colors.stripe),fill(colors.blue),fill(colors.note),fill(colors.available)].join('');
 const border='<border><left style="thin"><color rgb="FFDCE5EA"/></left><right style="thin"><color rgb="FFDCE5EA"/></right><top style="thin"><color rgb="FFDCE5EA"/></top><bottom style="thin"><color rgb="FFDCE5EA"/></bottom><diagonal/></border>';
 const xfs=[xf(0,0),xf(1,2,0,'left'),xf(2,2,0,'left'),xf(3,4),xf(4,4),xf(5,9,0,'left'),xf(6,3),xf(9,10,1,'left'),xf(9,10,1,'left'),xf(9,10,1),xf(9,10,1),xf(9,10,1),xf(9,10,1),xf(9,10,1),xf(9,10,1,'left')].join('');
 const dxf=(bg,fg)=>'<dxf><font><color rgb="'+fg+'"/></font><fill><patternFill patternType="solid"><fgColor rgb="'+bg+'"/><bgColor indexed="64"/></patternFill></fill></dxf>';
 return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="'+NS+'">'+
 '<fonts count="10">'+fonts+'</fonts><fills count="11">'+fills+'</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'+border+'</borders>'+
 '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="15">'+xfs+'</cellXfs>'+
 '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'+
 '<dxfs count="2">'+dxf('FF20252C','FFFFFFFF')+dxf('FFDFEEFF','FF1D568A')+'</dxfs></styleSheet>';
}
function worksheet(order,favorites,byName){
 const first=7,last=first+order.length-1;
 const meta=[
 '<row r="1" ht="39" customHeight="1">'+strCell('A1','TJRS 2026  |  ACOMPANHAMENTO DE VAGAS',1)+'</row>',
 '<row r="2" ht="25" customHeight="1">'+strCell('A2','Minha ordem de comarcas · quadro pessoal para preencher durante a escolha',2)+'</row>',
 '<row r="3" ht="26" customHeight="1">'+strCell('A3','VAGAS INICIAIS',3)+strCell('C3','ESCOLHIDAS (REGISTRADAS)',3)+strCell('E3','RESTANTES (CONFERIDAS)',3)+strCell('G3','SEM ATUALIZAÇÃO',3)+'</row>',
 '<row r="4" ht="42" customHeight="1">'+formulaCell('A4','SUM(C'+first+':C'+last+')',4)+formulaCell('C4','SUM(D'+first+':D'+last+')',4)+formulaCell('E4','SUM(E'+first+':E'+last+')',4)+formulaCell('G4','COUNTBLANK(D'+first+':D'+last+')',4)+'</row>',
 '<row r="5" ht="33" customHeight="1">'+strCell('A5','Todas começam DISPONÍVEIS. Informe 0 se nenhuma vaga foi escolhida; 1 ou mais para registrar ocupações. Azul = parcial · Preto = esgotada.',5)+'</row>',
 '<row r="6" ht="31" customHeight="1">'+
 ['Pos.','Comarca','Vagas totais','Escolhidas ✎','Restantes','Situação','UNICAA','+15%','Favorita'].map((t,i)=>strCell(String.fromCharCode(65+i)+'6',t,6)).join('')+'</row>'
 ];
 const rows=order.map((name,i)=>{
  const r=first+i,d=byName.get(name),alt=i%2===1,tx=alt?8:7,num=alt?10:9;
  return '<row r="'+r+'" ht="26" customHeight="1">'+
   numberCell('A'+r,i+1,num)+strCell('B'+r,name,(favorites.has(name)?14:tx))+
   numberCell('C'+r,d.vagas,num)+
   '<c r="D'+r+'" s="11"/>'+
   formulaCell('E'+r,'IF(D'+r+'="",C'+r+',MAX(0,C'+r+'-D'+r+'))',12)+
   formulaCell('F'+r,'IF(D'+r+'="","DISPONÍVEL",IF(D'+r+'>=C'+r+',"ESGOTADA",IF(D'+r+'>0,"PARCIAL","DISPONÍVEL")))',13,true)+
   numberCell('G'+r,d.unicaa,num)+strCell('H'+r,d.dificil?'Sim':'—',tx)+
   strCell('I'+r,favorites.has(name)?'★':'',num)+'</row>';
 }).join('');
 const columns=[8,32,14,17,15,20,13,11,12].map((w,i)=>'<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>').join('');
 const merge=['A1:I1','A2:I2','A3:B3','C3:D3','E3:F3','G3:I3','A4:B4','C4:D4','E4:F4','G4:I4','A5:I5'];
 // Use numeric inputs, not the calculated status text, to drive row colors.
 // Blank / zero = the green base style. Only positive numeric values trigger an override.
 const format='<conditionalFormatting sqref="A'+first+':I'+last+'">'+
 '<cfRule type="expression" dxfId="0" priority="1" stopIfTrue="1"><formula>AND(ISNUMBER($D'+first+'),$D'+first+'>=1,$D'+first+'>=$C'+first+')</formula></cfRule>'+
 '<cfRule type="expression" dxfId="1" priority="2" stopIfTrue="1"><formula>AND(ISNUMBER($D'+first+'),$D'+first+'>0,$D'+first+'&lt;$C'+first+')</formula></cfRule>'+
 '</conditionalFormatting>';
 return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="'+NS+'">'+
 '<dimension ref="A1:I'+last+'"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="D7" sqref="D7"/></sheetView></sheetViews>'+
 '<sheetFormatPr defaultRowHeight="20"/><cols>'+columns+'</cols><sheetData>'+meta.join('')+rows+'</sheetData>'+
 '<autoFilter ref="A6:I'+last+'"/>'+
 '<mergeCells count="'+merge.length+'">'+merge.map(ref=>'<mergeCell ref="'+ref+'"/>').join('')+'</mergeCells>'+format+
 '<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" errorTitle="Quantidade inválida" error="Informe um número entre 0 e o total de vagas." sqref="D'+first+':D'+last+'"><formula1>0</formula1><formula2>C'+first+'</formula2></dataValidation></dataValidations>'+
 '<printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.45" bottom="0.45" header="0.2" footer="0.2"/><pageSetup orientation="landscape" paperSize="9" fitToWidth="1"/>'+
 '</worksheet>';
}
window.exportComarcasXlsx=function({order,favorites,byName,download}){
 const entries=[
 ['[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'],
 ['_rels/.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'],
 ['docProps/core.xml','<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Acompanhamento pessoal de vagas TJRS 2026</dc:title><dc:creator>Escolha de Comarcas</dc:creator></cp:coreProperties>'],
 ['docProps/app.xml','<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>'],
 ['xl/workbook.xml','<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="'+NS+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Acompanhamento" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'],
 ['xl/_rels/workbook.xml.rels','<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
 ['xl/styles.xml',styles()],
 ['xl/worksheets/sheet1.xml',worksheet(order,favorites,byName)]
 ];
 download(zip(entries),'acompanhamento_comarcas_tjrs_2026.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
};
})();
