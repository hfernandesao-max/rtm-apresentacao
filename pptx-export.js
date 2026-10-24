/* Exportação da página "Integração RTM → SAP" para PowerPoint (PptxGenJS, no browser).
   Lê o conteúdo da página no idioma activo, por isso o ficheiro sai em PT ou EN. */
(function(){
var LIB='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
var C={navy:'273777',blue:'2397D4',green:'1D8A50',orange:'B45309',ink:'1D1D1F',muted:'5A6480',line:'D2D2D7',bg:'F5F7FB',tile:'FFFFFF',
  agt:'FDF1E6',rtm:'E8F2FB',int:'EEEEF6',sap:'E9F6EE'};
var F='Arial';
function en(){return (document.documentElement.lang||'').slice(0,2)==='en';}
function T(pt,e){return en()?e:pt;}
function txt(el){return el?el.textContent.replace(/\s+/g,' ').trim():'';}
function $(s,r){return (r||document).querySelector(s);}
function $$(s,r){return [].slice.call((r||document).querySelectorAll(s));}
function lines(el){ // texto com quebras <br> preservadas
  if(!el)return '';var c=el.cloneNode(true);$$('br',c).forEach(function(b){b.replaceWith('\n');});
  return c.textContent.split('\n').map(function(x){return x.replace(/\s+/g,' ').trim();}).filter(Boolean).join('\n');}
function load(cb){
  if(window.PptxGenJS)return cb();
  var s=document.createElement('script');s.src=LIB;s.onload=cb;
  s.onerror=function(){alert(T('Não foi possível carregar o gerador de PowerPoint. Verifique a ligação à internet.','Could not load the PowerPoint generator. Check the internet connection.'));busy(false);};
  document.head.appendChild(s);
}
var btn;
function busy(on){if(!btn)return;btn.disabled=on;btn.querySelector('span').textContent=on?T('A criar…','Creating…'):T('PowerPoint','PowerPoint');}

function build(){
  var p=new PptxGenJS();p.layout='LAYOUT_WIDE';p.title=document.title;p.company='Grupo Naval Angola';
  var W=13.333,H=7.5,M=0.5,n=0;
  var FOOT=T('Integração RTM → SAP · Grupo Naval Angola · Projecto RTM · Confidencial','RTM → SAP integration · Grupo Naval Angola · RTM Project · Confidential');
  function slide(kicker,title){
    var s=p.addSlide();n++;s.background={color:C.bg};
    s.addShape(p.ShapeType.rect,{x:0,y:0,w:W,h:0.12,fill:{color:C.navy},line:{color:C.navy}});
    if(kicker)s.addText(kicker,{x:M,y:0.32,w:W-2*M,h:0.3,fontFace:F,fontSize:12,bold:true,color:C.blue});
    if(title)s.addText(title,{x:M,y:0.6,w:W-2*M,h:0.7,fontFace:F,fontSize:26,bold:true,color:C.ink,valign:'top',fit:'shrink'});
    s.addText(FOOT,{x:M,y:H-0.42,w:W-2*M-1,h:0.3,fontFace:F,fontSize:9,color:C.muted});
    s.addText(String(n),{x:W-M-0.8,y:H-0.42,w:0.8,h:0.3,fontFace:F,fontSize:9,color:C.muted,align:'right'});
    return s;
  }
  function head(sec){return {k:txt($('.kicker',sec)),t:txt($('h2',sec)||$('h1',sec))};}
  function card(s,x,y,w,h,title,items,accent){
    s.addShape(p.ShapeType.roundRect,{x:x,y:y,w:w,h:h,rectRadius:0.12,fill:{color:C.tile},line:{color:C.line,width:0.75}});
    if(accent)s.addShape(p.ShapeType.rect,{x:x,y:y,w:w,h:0.07,fill:{color:accent},line:{color:accent}});
    s.addText(title,{x:x+0.2,y:y+0.15,w:w-0.4,h:0.45,fontFace:F,fontSize:14,bold:true,color:C.ink,valign:'top',fit:'shrink'});
    var body=items.map(function(t){return {text:t,options:{bullet:{indent:12},breakLine:true}};});
    s.addText(body,{x:x+0.2,y:y+0.62,w:w-0.4,h:h-0.75,fontFace:F,fontSize:11,color:C.muted,valign:'top',paraSpaceAfter:4,fit:'shrink'});
  }

  // 1 · Capa
  var hero=$$('section')[0];
  var s=p.addSlide();n++;s.background={color:C.navy};
  s.addText(txt($('.kicker',hero)),{x:M,y:1.2,w:W-2*M,h:0.4,fontFace:F,fontSize:14,bold:true,color:'9FB3E8'});
  s.addText(lines($('h1',hero)),{x:M,y:1.7,w:W-2*M,h:2.0,fontFace:F,fontSize:40,bold:true,color:'FFFFFF',valign:'top',fit:'shrink'});
  s.addText(txt($('.lead',hero)),{x:M,y:3.9,w:W-2*M-1,h:1.4,fontFace:F,fontSize:16,color:'DCE3F5',valign:'top',fit:'shrink'});
  s.addText('Grupo Naval Angola · '+T('Projecto RTM · Confidencial','RTM Project · Confidential'),{x:M,y:H-0.9,w:W-2*M,h:0.4,fontFace:F,fontSize:12,color:'9FB3E8'});

  // 2 · Princípios
  s=slide(txt($('.kicker',hero)),T('Os três princípios da integração','The three principles of the integration'));
  var rules=$$('.rule > div'),cw=(W-2*M-0.5)/3,acc=[C.blue,C.green,C.orange];
  rules.forEach(function(r,i){
    var ps=$$('p',r).map(txt);
    card(s,M+i*(cw+0.25),1.6,cw,4.9,txt($('h3',r)),ps,acc[i]);
  });

  // 3 · Documentos
  var sec=$('#documentos'),h=head(sec);s=slide(h.k,h.t);
  var rows=[[T('Doc.','Doc.'),T('Documento','Document'),'AGT',T('O que cria no outro sistema','What it creates in the other system')].map(function(t){return {text:t,options:{bold:true,color:'FFFFFF',fill:{color:C.navy}}};})];
  $$('.doc',sec).forEach(function(d){
    rows.push([{text:txt($('.code',d)),options:{bold:true,color:d.classList.contains('sapdoc')?C.green:C.navy}},
      {text:txt($('h3',d))+'\n'+txt($('.ds',d)),options:{}},
      {text:txt($('.tag',d)),options:{color:C.muted}},
      {text:$$('li',d).map(txt).join('\n'),options:{}}]);
  });
  s.addTable(rows,{x:M,y:1.45,w:W-2*M,colW:[0.7,2.6,2.1,W-2*M-5.4],fontFace:F,fontSize:9.5,color:C.ink,border:{type:'solid',pt:0.5,color:C.line},fill:{color:C.tile},valign:'middle',autoPage:true,autoPageRepeatHeader:true,newSlideStartY:1.45});

  // 4 · Reposição automática
  sec=$('#reposicao');h=head(sec);s=slide(h.k,h.t);
  var cys=$$('.cy',sec),gap=0.18,bw=(W-2*M-gap*(cys.length-1))/cys.length;
  cys.forEach(function(c,i){
    var x=M+i*(bw+gap);
    s.addShape(p.ShapeType.roundRect,{x:x,y:1.55,w:bw,h:2.6,rectRadius:0.1,fill:{color:C.tile},line:{color:C.line,width:0.75}});
    s.addText(String(i+1),{x:x+0.15,y:1.68,w:0.4,h:0.4,shape:p.ShapeType.ellipse,fill:{color:C.blue},color:'FFFFFF',bold:true,fontSize:12,align:'center',valign:'middle',fontFace:F});
    s.addText(txt($('.cy-w',c)).toUpperCase(),{x:x+0.15,y:2.15,w:bw-0.3,h:0.3,fontFace:F,fontSize:8.5,bold:true,color:C.blue,fit:'shrink'});
    s.addText(txt($('h3',c)),{x:x+0.15,y:2.42,w:bw-0.3,h:0.4,fontFace:F,fontSize:12.5,bold:true,color:C.ink,fit:'shrink'});
    s.addText(txt($('.cy-t',c)),{x:x+0.15,y:2.85,w:bw-0.3,h:1.25,fontFace:F,fontSize:9.5,color:C.muted,valign:'top',fit:'shrink'});
    if(i<cys.length-1)s.addShape(p.ShapeType.rightArrow,{x:x+bw+0.01,y:2.75,w:gap-0.02,h:0.22,fill:{color:C.blue},line:{color:C.blue}});
  });
  var cards=$$('.card',sec),cw2=(W-2*M-0.3)/2;
  cards.forEach(function(c,i){card(s,M+i*(cw2+0.3),4.4,cw2,2.5,txt($('h3',c)),$$('li',c).map(txt),i?C.orange:C.blue);});

  // 5 · Visão geral por raias
  sec=$('#fluxo');h=head(sec);s=slide(h.k,h.t);
  var cells=[].slice.call($('.lanes',sec).children),hdr=cells.slice(0,7),grid=[];
  grid.push(hdr.map(function(c){return {text:txt(c),options:{bold:true,color:'FFFFFF',fill:{color:C.navy},align:'center'}};}));
  var laneFill=[C.agt,C.rtm,C.int,C.sap];
  for(var r=0;r<4;r++){
    var row=cells.slice(7+r*7,14+r*7);
    grid.push(row.map(function(c,j){
      if(j===0)return {text:txt(c),options:{bold:true,fill:{color:laneFill[r]}}};
      var b=txt($('b',c)),sp=txt($('span:not(.dn)',c));
      return {text:b?(b+(sp?'\n'+sp:'')):'',options:{fill:{color:b?laneFill[r]:C.tile}}};
    }));
  }
  s.addTable(grid,{x:M,y:1.45,w:W-2*M,colW:[1.3].concat(Array(6).fill((W-2*M-1.3)/6)),fontFace:F,fontSize:9,color:C.ink,border:{type:'solid',pt:0.5,color:C.line},valign:'middle',rowH:[0.45,1.1,1.1,1.1,1.1]});

  // 6–8 · Detalhe por separador
  sec=$('#detalhe');var tabs=$$('.tabs button',sec);
  tabs.forEach(function(b){
    var panel=document.getElementById(b.getAttribute('data-p'));
    s=slide(txt($('.kicker',sec)),txt($('h2',sec))+' · '+txt(b).replace(/^[^\wÀ-ÿ]+/,''));
    var t=$('table',panel),rows=[];
    rows.push($$('th',t).map(function(th){return {text:txt(th),options:{bold:true,color:'FFFFFF',fill:{color:C.navy}}};}));
    $$('tbody tr',t).forEach(function(tr){
      rows.push($$('td',tr).map(function(td,j){
        var main=$$('b, .tag',td).map(txt).join(' · ')||'',sm=txt($('small',td));
        if(!main)main=txt(td);
        return {text:main+(sm?'\n'+sm:''),options:j===0?{bold:false}:{}};
      }));
    });
    s.addTable(rows,{x:M,y:1.45,w:W-2*M,colW:[2.6,2.2,4.7,W-2*M-9.5],fontFace:F,fontSize:9,color:C.ink,border:{type:'solid',pt:0.5,color:C.line},fill:{color:C.tile},valign:'top',autoPage:true,autoPageRepeatHeader:true,newSlideStartY:1.45});
    var note=$('.note',panel);
    if(note&&rows.length<7)s.addText(txt(note),{x:M,y:6.35,w:W-2*M,h:0.55,fontFace:F,fontSize:10,italic:true,color:C.muted,fit:'shrink'});
  });

  // 9 · Controlos
  sec=$('#controlos');h=head(sec);s=slide(h.k,h.t);
  var cc=$$('.card',sec),cols=3,cw3=(W-2*M-0.5)/cols,ch=(H-1.55-0.6-0.25)/2;
  cc.forEach(function(c,i){var col=i%cols,row=Math.floor(i/cols);card(s,M+col*(cw3+0.25),1.5+row*(ch+0.2),cw3,ch,txt($('h3',c)),$$('li',c).map(txt),C.blue);});

  // 10 · Um dia de operação
  sec=$('#dia');h=head(sec);s=slide(h.k,h.t);
  var ds=$$('.day > div',sec),half=Math.ceil(ds.length/2);
  ds.forEach(function(d,i){
    var col=i<half?0:1,row=i<half?i:i-half,x=M+col*((W-2*M)/2),y=1.55+row*1.12;
    s.addShape(p.ShapeType.ellipse,{x:x,y:y+0.06,w:0.2,h:0.2,fill:{color:C.tile},line:{color:C.blue,width:2}});
    s.addText(txt($('b',d)),{x:x+0.35,y:y,w:5.6,h:0.3,fontFace:F,fontSize:10.5,bold:true,color:C.blue});
    s.addText(txt($('p',d)),{x:x+0.35,y:y+0.28,w:5.6,h:0.32,fontFace:F,fontSize:12.5,bold:true,color:C.ink,fit:'shrink'});
    s.addText(txt($('small',d)),{x:x+0.35,y:y+0.6,w:5.6,h:0.45,fontFace:F,fontSize:10,color:C.muted,valign:'top',fit:'shrink'});
  });

  // 11 · Alternativa
  var alt=$('.cmp').closest('section');h=head(alt);s=slide(h.k,h.t);
  $$('.cmp .card',alt).forEach(function(c,i){
    card(s,M+i*(cw2+0.3),1.55,cw2,4.9,txt($('.k',c))+' · '+txt($('h3',c)),$$('li',c).map(txt),i?C.muted:C.green);
  });

  // 12 · Pontos a validar
  sec=$('#validar');h=head(sec);s=slide(h.k,h.t);
  var vc=$$('.card',sec),cw4=(W-2*M-0.3)/2,ch4=(H-1.55-0.6-0.2)/2;
  vc.forEach(function(c,i){var col=i%2,row=Math.floor(i/2);card(s,M+col*(cw4+0.3),1.5+row*(ch4+0.2),cw4,ch4,txt($('h3',c)),$$('li',c).map(txt),C.orange);});

  var d=new Date(),stamp=d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);
  return p.writeFile({fileName:T('Integracao_RTM_SAP_','RTM_SAP_Integration_')+stamp+(en()?'_EN':'_PT')+'.pptx'});
}

document.addEventListener('DOMContentLoaded',init);if(document.readyState!=='loading')init();
function init(){
  if(btn)return;btn=document.getElementById('pptx-btn');if(!btn)return;
  btn.addEventListener('click',function(){busy(true);load(function(){
    Promise.resolve().then(build).then(function(){busy(false);}).catch(function(e){console.error(e);alert(T('Não foi possível criar o PowerPoint.','Could not create the PowerPoint.'));busy(false);});
  });});
  new MutationObserver(function(){if(!btn.disabled)busy(false);}).observe(document.documentElement,{attributes:true,attributeFilter:['lang']});
}
})();
