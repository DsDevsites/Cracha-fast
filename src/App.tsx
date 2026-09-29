import {useEffect,useMemo,useRef,useState} from 'react'
import {BadgeCheck,FileDown,FileSpreadsheet,LayoutDashboard,Palette,Plus,Printer,Settings,Trash2,Upload,Users,Search,Save,Sparkles,ImagePlus,ScanText,MousePointer2,RotateCcw} from 'lucide-react'
import {createWorker} from 'tesseract.js'
import type {BadgeTemplate,Batch,FieldKey,FieldPosition,Person} from './types'
import {defaultTemplate,loadBatches,loadPeople,loadTemplates,saveBatches,savePeople,saveTemplates} from './lib/storage'
import {importSpreadsheet} from './lib/importers'
import {exportDocx} from './lib/docxExport'

type Tab='dashboard'|'people'|'templates'|'generate'|'settings'
const emptyPerson:Person={id:'preview',name:'Nome do colaborador',registration:'00000',role:'Função'}
const defaultFields:Record<FieldKey,FieldPosition>={name:{x:8,y:67,width:84,height:7,fontSize:7,fontWeight:'normal',align:'left'},registration:{x:8,y:75,width:84,height:7,fontSize:7,fontWeight:'normal',align:'left'},role:{x:8,y:83,width:84,height:7,fontSize:7,fontWeight:'normal',align:'left'}}

function normalizedTemplate(t:BadgeTemplate):BadgeTemplate{const fields=t.fieldPositions||defaultFields;const normalized=Object.fromEntries((Object.keys(defaultFields) as FieldKey[]).map(k=>[k,{...defaultFields[k],...(fields[k]||{})}])) as Record<FieldKey,FieldPosition>;return{...t,fieldPositions:normalized}}
function Badge({person,template}:{person:Person;template:BadgeTemplate}){
  const t=normalizedTemplate(template); const f=t.fieldPositions!
  if(t.backgroundDataUrl)return <div className="image-badge">
    <img className="image-badge-bg" src={t.backgroundDataUrl} alt="" />
    <div className="image-badge-field" style={fieldStyle(f.name)}>{person.name||'NOME'}</div>
    <div className="image-badge-field" style={fieldStyle(f.registration)}>{person.registration||'MATRÍCULA'}</div>
    <div className="image-badge-field" style={fieldStyle(f.role)}>{person.role||'FUNÇÃO'}</div>
  </div>
  return <div className="badge"><div className="badge-symbol"><div className="symbol-triangle">⚠</div></div><div className="badge-title" style={{background:t.primary}}>{t.title}<br/>{t.subtitle}</div><div className="badge-fields"><div><b>NOME:</b> {person.name||'—'}</div><div><b>MATRÍCULA:</b> {person.registration||'—'}</div><div><b>FUNÇÃO:</b> {person.role||'—'}</div></div><div className="badge-logo"><span className="logo-mark">◎</span><strong style={{color:t.primary}}>{t.logoText}</strong></div></div>
}
function fieldStyle(p:FieldPosition){return{left:p.x+'%',top:p.y+'%',width:p.width+'%',height:p.height+'%',fontSize:p.fontSize+'pt',fontWeight:p.fontWeight,textAlign:p.align}}
function Nav({active,icon,text,onClick}:{active:boolean;icon:React.ReactNode;text:string;onClick:()=>void}){return <button className={active?'nav active':'nav'} onClick={onClick}>{icon}<span>{text}</span></button>}
function Stat({label,value,icon}:{label:string;value:string|number;icon:React.ReactNode}){return <div className="stat"><div className="stat-icon">{icon}</div><div><span>{label}</span><strong>{value}</strong></div></div>}

function VisualEditor({template,onChange}:{template:BadgeTemplate;onChange:(p:Partial<BadgeTemplate>)=>void}){
  const t=normalizedTemplate(template)
  const [selected,setSelected]=useState<FieldKey>('name')
  const [busy,setBusy]=useState(false)
  const inputRef=useRef<HTMLInputElement>(null)

  const updateField=(key:FieldKey,p:Partial<FieldPosition>)=>
    onChange({fieldPositions:{...t.fieldPositions!,[key]:{...t.fieldPositions![key],...p}}})

  const loadImage=(file?:File)=>{
    if(!file)return
    const reader=new FileReader()
    reader.onload=()=>{
      const img=new Image()
      img.onload=()=>{
        const max=1400
        const scale=Math.min(1,max/img.width,max/img.height)
        const canvas=document.createElement('canvas')
        canvas.width=Math.max(1,Math.round(img.width*scale))
        canvas.height=Math.max(1,Math.round(img.height*scale))
        canvas.getContext('2d')?.drawImage(img,0,0,canvas.width,canvas.height)
        onChange({backgroundDataUrl:canvas.toDataURL('image/jpeg',0.92),fieldPositions:defaultFields})
      }
      img.src=String(reader.result)
    }
    reader.readAsDataURL(file)
  }

  const paste=(e:React.ClipboardEvent<HTMLDivElement>)=>{
    const item=[...e.clipboardData.items].find(x=>x.type.startsWith('image/'))
    if(item){e.preventDefault();loadImage(item.getAsFile()||undefined)}
  }

  const autoDetect=async()=>{
    if(!t.backgroundDataUrl)return
    setBusy(true)
    try{
      const img=new Image()
      img.src=t.backgroundDataUrl
      await new Promise<void>((resolve,reject)=>{img.onload=()=>resolve();img.onerror=reject})
      const imageWidth=img.naturalWidth||1000
      const imageHeight=img.naturalHeight||1000
      const worker=await createWorker('por')
      const result=await worker.recognize(t.backgroundDataUrl)
      const words=(result as any).data.words||[]
      const next={...t.fieldPositions!}
      for(const w of words){
        const s=String(w.text||'').toUpperCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
        const b=w.bbox
        if(!b)continue
        let key:FieldKey|undefined
        if(s.includes('NOME'))key='name'
        else if(s.includes('MATRICULA')||s.includes('MATR'))key='registration'
        else if(s.includes('FUNCAO')||s.includes('FUNCA'))key='role'
        if(key){
          next[key]={
            ...next[key],
            x:Math.max(0,Math.min(90,(b.x0/imageWidth)*100)),
            y:Math.max(0,Math.min(92,(b.y1/imageHeight)*100+1)),
            width:Math.min(90,Math.max(20,(b.x1-b.x0)/imageWidth*100+30)),
            height:8
          }
        }
      }
      await worker.terminate()
      onChange({fieldPositions:next})
    }catch{
      alert('Não foi possível reconhecer os campos automaticamente. Ajuste-os manualmente.')
    }finally{setBusy(false)}
  }

  return <div className="visual-editor" onPaste={paste} tabIndex={0}>
    <div className="editor-toolbar">
      <div><b>Montar modelo</b><span>Cole a foto do crachá com Ctrl+V ou envie uma imagem.</span></div>
      <div className="toolbar-actions">
        <button className="secondary" type="button" onClick={()=>inputRef.current?.click()}><ImagePlus size={16}/> Imagem</button>
        <button className="secondary" type="button" disabled={!t.backgroundDataUrl||busy} onClick={autoDetect}><ScanText size={16}/> {busy?'Reconhecendo...':'Reconhecer campos'}</button>
        <button className="secondary" type="button" disabled={!t.backgroundDataUrl} onClick={()=>onChange({backgroundDataUrl:undefined,fieldPositions:defaultFields})}><RotateCcw size={16}/> Modelo padrão</button>
      </div>
      <input ref={inputRef} hidden type="file" accept="image/*" onChange={e=>loadImage(e.target.files?.[0])}/>
    </div>

    <div className="editor-help"><MousePointer2 size={15}/> Clique e arraste as caixas NOME, MATRÍCULA e FUNÇÃO para o lugar exato.</div>

    <div className="editor-workspace">
      <div className="canvas-wrap">
        {t.backgroundDataUrl
          ? <div className="template-canvas" style={{backgroundImage:`url(${t.backgroundDataUrl})`}}>
              {(['name','registration','role'] as FieldKey[]).map(k=>
                <div key={k} className={'drag-field '+(selected===k?'selected':'')} style={fieldStyle(t.fieldPositions![k])}
                  onMouseDown={e=>{
                    e.preventDefault()
                    setSelected(k)
                    const startX=e.clientX,startY=e.clientY,start={...t.fieldPositions![k]}
                    const rect=(e.currentTarget.parentElement as HTMLElement).getBoundingClientRect()
                    const move=(ev:MouseEvent)=>{
                      updateField(k,{
                        x:Math.max(0,Math.min(100-start.width,start.x+(ev.clientX-startX)/rect.width*100)),
                        y:Math.max(0,Math.min(100-start.height,start.y+(ev.clientY-startY)/rect.height*100))
                      })
                    }
                    const up=()=>{window.removeEventListener('mousemove',move);window.removeEventListener('mouseup',up)}
                    window.addEventListener('mousemove',move);window.addEventListener('mouseup',up)
                  }}>
                  {k==='name'?'NOME: '+(emptyPerson.name):k==='registration'?'MATRÍCULA: '+(emptyPerson.registration):'FUNÇÃO: '+(emptyPerson.role)}<small>arraste</small>
                </div>)}
            </div>
          : <div className="drop-zone">
              <ImagePlus size={42}/>
              <b>Cole ou envie o crachá pronto</b>
              <span>Ctrl+V dentro desta área ou clique em “Imagem”.</span>
              <button className="primary" type="button" onClick={()=>inputRef.current?.click()}><Upload size={16}/> Escolher imagem</button>
            </div>}
      </div>

      <div className="field-controls">
        <h4>Campo selecionado</h4>
        <select value={selected} onChange={e=>setSelected(e.target.value as FieldKey)}>
          <option value="name">Nome</option><option value="registration">Matrícula</option><option value="role">Função</option>
        </select>
        {(['x','y','width','height'] as const).map(k=>
          <label key={k}>{k==='x'?'X':k==='y'?'Y':k==='width'?'Largura':'Altura'}
            <input type="number" min="0" max="100" step=".5" value={Number(t.fieldPositions![selected][k])}
              onChange={e=>updateField(selected,{[k]:Number(e.target.value)})}/>
          </label>)}
        <div className="font-preview-label">Tamanho da letra: <b>{t.fieldPositions![selected].fontSize} pt</b></div>
        <label>Tamanho da letra
          <input type="number" min="4" max="30" step=".5" value={t.fieldPositions![selected].fontSize}
            onChange={e=>updateField(selected,{fontSize:Number(e.target.value)})}/>
        </label>
        <label>Peso da letra
          <select value={t.fieldPositions![selected].fontWeight} onChange={e=>updateField(selected,{fontWeight:e.target.value as 'normal'|'bold'})}>
            <option value="normal">Normal</option><option value="bold">Negrito</option>
          </select>
        </label>
        <label>Alinhamento
          <select value={t.fieldPositions![selected].align} onChange={e=>updateField(selected,{align:e.target.value as 'left'|'center'|'right'})}>
            <option value="left">Esquerda</option><option value="center">Centro</option><option value="right">Direita</option>
          </select>
        </label>
        <div className="font-example" style={{fontSize:t.fieldPositions![selected].fontSize+'pt',fontWeight:t.fieldPositions![selected].fontWeight,textAlign:t.fieldPositions![selected].align}}>
          {selected==='name'?'ALBERONI MACEDO DE SENA':selected==='registration'?'123456':'TÉCNICO DE SEGURANÇA'}
        </div>
        <div className="field-tip">O OCR procura NOME, MATRÍCULA e FUNÇÃO. Se não encontrar algum campo, basta posicioná-lo manualmente.</div>
      </div>
    </div>
  </div>
}
function App(){
 const[tab,setTab]=useState<Tab>('dashboard');const[people,setPeople]=useState<Person[]>(loadPeople);const[templates,setTemplates]=useState<BadgeTemplate[]>(loadTemplates().map(normalizedTemplate));const[batches,setBatches]=useState<Batch[]>(loadBatches());const[selectedTemplate,setSelectedTemplate]=useState(defaultTemplate.id);const[search,setSearch]=useState('');const[fileRef]=useState(()=>({current:null as HTMLInputElement|null}));const template=normalizedTemplate(templates.find(t=>t.id===selectedTemplate)||templates[0]);const filtered=useMemo(()=>people.filter(p=>(p.name+' '+p.registration+' '+p.role).toLowerCase().includes(search.toLowerCase())),[people,search]);const updatePeople=(v:Person[])=>{setPeople(v);savePeople(v)};const updateTemplates=(v:BadgeTemplate[])=>{setTemplates(v);saveTemplates(v)};const edit=(p:Partial<BadgeTemplate>)=>updateTemplates(templates.map(t=>t.id===template.id?normalizedTemplate({...t,...p}):t));const importFile=async(f?:File)=>{if(!f)return;try{const x=await importSpreadsheet(f);updatePeople([...people,...x]);setTab('people');alert(x.length+' pessoas importadas.')}catch{alert('Planilha inválida. Use XLSX, XLS ou CSV.')}};const gen=()=>{if(!people.length)return alert('Cadastre ou importe pessoas primeiro.');const b:Batch={id:crypto.randomUUID(),name:'Lote '+new Date().toLocaleString('pt-BR'),templateId:template.id,people,createdAt:new Date().toISOString()};const n=[b,...batches];setBatches(n);saveBatches(n);setTab('generate')};useEffect(()=>{const fn=(e:ClipboardEvent)=>{if(tab!=='templates'||!e.clipboardData)return;const item=[...e.clipboardData.items].find(x=>x.type.startsWith('image/'));if(item){const el=document.querySelector('.visual-editor') as HTMLElement|null;el?.focus()}};window.addEventListener('paste',fn);return()=>window.removeEventListener('paste',fn)},[tab])
 return <div className="app"><aside className="sidebar"><div className="brand"><div className="brand-icon"><BadgeCheck size={22}/></div><div><strong>Cracha Fast</strong><small>Carteirinhas em lote</small></div></div><nav><Nav active={tab==='dashboard'} icon={<LayoutDashboard/>} text="Dashboard" onClick={()=>setTab('dashboard')}/><Nav active={tab==='people'} icon={<Users/>} text="Pessoas" onClick={()=>setTab('people')}/><Nav active={tab==='templates'} icon={<Palette/>} text="Modelos" onClick={()=>setTab('templates')}/><Nav active={tab==='generate'} icon={<Printer/>} text="Gerar crachás" onClick={()=>setTab('generate')}/><Nav active={tab==='settings'} icon={<Settings/>} text="Configurações" onClick={()=>setTab('settings')}/></nav><div className="side-bottom"><div className="mini-plan"><Sparkles size={17}/><div><b>Plano Demo</b><small>Pronto para evoluir</small></div></div></div></aside><main className="main"><header className="topbar"><div><span className="eyebrow">GERADOR DE CARTEIRINHAS</span><h1>{tab==='dashboard'?'Visão geral':tab==='people'?'Pessoas':tab==='templates'?'Modelos de crachá':tab==='generate'?'Gerar crachás':'Configurações'}</h1></div><button className="primary" onClick={()=>setTab('generate')}><Printer size={17}/> Gerar</button></header>
 {tab==='dashboard'&&<section className="page"><div className="hero"><div><span className="pill">12 por folha A4</span><h2>Crie lotes de carteirinhas em poucos cliques.</h2><p>Agora você também pode usar uma foto do crachá pronto como fundo e posicionar Nome, Matrícula e Função por cima.</p><button className="primary" onClick={()=>setTab('templates')}><Palette size={17}/> Montar modelo</button></div><div className="hero-preview"><Badge person={people[0]||emptyPerson} template={template}/></div></div><div className="stats"><Stat label="Pessoas" value={people.length} icon={<Users/>}/><Stat label="Modelos" value={templates.length} icon={<Palette/>}/><Stat label="Lotes gerados" value={batches.length} icon={<FileDown/>}/><Stat label="Por folha" value="12" icon={<FileSpreadsheet/>}/></div><div className="quick-grid"><button className="quick" onClick={()=>setTab('templates')}><span>01</span><div><b>Monte o modelo</b><p>Envie ou cole a foto do crachá pronto.</p></div><Plus size={18}/></button><button className="quick" onClick={()=>setTab('people')}><span>02</span><div><b>Importe a lista</b><p>Nome, matrícula e função em XLSX/CSV.</p></div><Plus size={18}/></button><button className="quick" onClick={gen}><span>03</span><div><b>Gere em lote</b><p>12 por A4, impressão ou Word.</p></div><Plus size={18}/></button></div></section>}
 {tab==='people'&&<section className="page"><div className="toolbar"><div className="search"><Search size={17}/><input placeholder="Buscar por nome, matrícula ou função..." value={search} onChange={e=>setSearch(e.target.value)}/></div><div className="toolbar-actions"><button className="secondary" onClick={()=>document.getElementById('people-file')?.click()}><Upload size={17}/> Importar</button><button className="primary" onClick={()=>updatePeople([...people,{id:crypto.randomUUID(),name:'',registration:'',role:''}])}><Plus size={17}/> Adicionar</button></div></div><input id="people-file" type="file" hidden accept=".xlsx,.xls,.csv" onChange={e=>importFile(e.target.files?.[0])}/><div className="table-wrap"><table><thead><tr><th>Nome</th><th>Matrícula</th><th>Função</th><th></th></tr></thead><tbody>{filtered.map(p=><tr key={p.id}><td><input value={p.name} onChange={e=>updatePeople(people.map(x=>x.id===p.id?{...x,name:e.target.value}:x))}/></td><td><input value={p.registration} onChange={e=>updatePeople(people.map(x=>x.id===p.id?{...x,registration:e.target.value}:x))}/></td><td><input value={p.role} onChange={e=>updatePeople(people.map(x=>x.id===p.id?{...x,role:e.target.value}:x))}/></td><td><button className="icon-btn danger" onClick={()=>updatePeople(people.filter(x=>x.id!==p.id))}><Trash2 size={16}/></button></td></tr>)}</tbody></table>{!filtered.length&&<div className="empty"><Users size={34}/><b>Nenhuma pessoa cadastrada</b><span>Importe uma planilha ou adicione manualmente.</span></div>}</div></section>}
 {tab==='templates'&&<section className="page"><div className="template-layout"><div className="panel"><div className="section-head"><div><h3>Editor de modelo</h3><p>Você pode criar do zero ou usar um crachá pronto.</p></div><button className="secondary" onClick={()=>{const t=normalizedTemplate({...defaultTemplate,id:crypto.randomUUID(),name:'Novo modelo'});updateTemplates([...templates,t]);setSelectedTemplate(t.id)}}><Plus size={17}/> Novo modelo</button></div><label>Modelo<select value={selectedTemplate} onChange={e=>setSelectedTemplate(e.target.value)}>{templates.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}</select></label><label>Nome do modelo<input value={template.name} onChange={e=>edit({name:e.target.value})}/></label><label>Cor principal<input type="color" value={template.primary} onChange={e=>edit({primary:e.target.value})}/></label><button className="primary full" onClick={()=>alert('Modelo salvo automaticamente.') }><Save size={17}/> Salvar modelo</button></div><VisualEditor template={template} onChange={edit}/></div></section>}
 {tab==='generate'&&<section className="page"><div className="generate-panel"><div><span className="eyebrow">MODELO DE IMPRESSÃO</span><h2>Escolha o modelo que deseja gerar</h2><p>{people.length} pessoas serão preenchidas com os dados cadastrados.</p></div><label className="generate-select-label">Modelo<select value={selectedTemplate} onChange={e=>setSelectedTemplate(e.target.value)}>{templates.map(t=><option key={t.id} value={t.id}>{t.name}</option>)}</select></label><div className="generate-preview-mini"><Badge person={people[0]||emptyPerson} template={template}/></div><div className="toolbar-actions"><button className="secondary" onClick={()=>window.print()}><Printer size={17}/> Imprimir / PDF</button><button className="primary" onClick={()=>exportDocx(people,template)}><FileDown size={17}/> Word editável</button></div></div><div className="generate-info"><b>{template.name}</b><span>{template.cardsPerPage} por folha A4 · {template.backgroundDataUrl?'modelo com imagem':'modelo criado no sistema'}</span></div><div className="print-sheet">{people.map(p=><Badge key={p.id} person={p} template={template}/>)}</div><div className="section-head"><div><h3>Histórico</h3><p>Os lotes ficam registrados neste navegador.</p></div></div><div className="history">{batches.slice(0,8).map(b=><div className="history-row" key={b.id}><div><b>{b.name}</b><span>{b.people.length} carteirinhas · {new Date(b.createdAt).toLocaleString('pt-BR')}</span></div><button className="secondary" onClick={()=>exportDocx(b.people,templates.find(t=>t.id===b.templateId)||template)}>Baixar Word</button></div>)}</div></section>}
 {tab==='settings'&&<section className="page"><div className="panel narrow"><div className="section-head"><div><h3>Configurações</h3><p>Preparação para a versão SaaS multiempresa.</p></div></div><div className="settings-card"><BadgeCheck/><div><b>Editor visual ativo</b><span>Modelos podem usar uma imagem pronta como fundo. Os campos são posicionados separadamente e ficam preenchidos com os dados da lista.</span></div></div><button className="secondary" onClick={()=>{localStorage.clear();location.reload()}}><Trash2 size={17}/> Limpar dados locais</button></div></section>}</main></div>
}
export default App