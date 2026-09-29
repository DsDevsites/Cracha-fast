import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, ImageRun, VerticalAlign } from 'docx'
import type { BadgeTemplate, FieldKey, Person } from '../types'

const fields:FieldKey[]=['name','registration','role']

function normalized(t:BadgeTemplate){
  const d:any={
    name:{x:8,y:67,width:84,height:7,fontSize:7,fontWeight:'normal',align:'left'},
    registration:{x:8,y:75,width:84,height:7,fontSize:7,fontWeight:'normal',align:'left'},
    role:{x:8,y:83,width:84,height:7,fontSize:7,fontWeight:'normal',align:'left'}
  }
  return fields.reduce((o,k)=>({...o,[k]:{...d[k],...(t.fieldPositions?.[k]||{})}}),{} as Record<FieldKey,any>)
}

function dataUrlToBytes(dataUrl:string){
  const base64=dataUrl.split(',')[1]||''
  const bin=atob(base64)
  const bytes=new Uint8Array(bin.length)
  for(let i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i)
  return bytes
}

async function renderBadge(t:BadgeTemplate,p:Person):Promise<Uint8Array>{
  const width=480,height=740
  const canvas=document.createElement('canvas')
  canvas.width=width;canvas.height=height
  const ctx=canvas.getContext('2d')!
  ctx.clearRect(0,0,width,height)
  if(t.backgroundDataUrl){
    const img=new Image()
    img.src=t.backgroundDataUrl
    await new Promise<void>((resolve,reject)=>{img.onload=()=>resolve();img.onerror=reject})
    ctx.drawImage(img,0,0,width,height)
  }else{
    ctx.fillStyle='#fff';ctx.fillRect(0,0,width,height)
    ctx.fillStyle=t.primary;ctx.fillRect(0,250,width,110)
    ctx.fillStyle='#fff';ctx.font='bold 28px Arial';ctx.textAlign='center'
    ctx.fillText(t.title,240,295);ctx.fillText(t.subtitle,240,330)
  }
  const f=normalized(t)
  const values:Record<FieldKey,string>={name:p.name||'NOME',registration:p.registration||'MATRÍCULA',role:p.role||'FUNÇÃO'}
  ctx.textBaseline='middle'
  for(const key of fields){
    const box=f[key]
    const x=box.x/100*width
    const y=box.y/100*height
    const w=box.width/100*width
    const h=box.height/100*height
    const px=Math.max(4,box.fontSize*96/72*(width/170))
    ctx.save()
    ctx.beginPath();ctx.rect(x,y,w,h);ctx.clip()
    ctx.fillStyle='#111'
    ctx.font=(box.fontWeight==='bold'?'700 ':'400 ')+px+'px Arial'
    ctx.textAlign=box.align==='center'?'center':box.align==='right'?'right':'left'
    const tx=box.align==='center'?x+w/2:box.align==='right'?x+w:x+2
    ctx.fillText(values[key],tx,y+h/2)
    ctx.restore()
  }
  const data=canvas.toDataURL('image/png')
  return dataUrlToBytes(data)
}

export async function exportDocx(people:Person[],t:BadgeTemplate){
  const rows:TableRow[]=[]
  for(let i=0;i<people.length;i+=4){
    const cells:TableCell[]=[]
    for(let j=0;j<4;j++){
      const p=people[i+j]
      if(!p){cells.push(new TableCell({children:[new Paragraph('')]}));continue}
      const image=await renderBadge(t,p)
      cells.push(new TableCell({
        verticalAlign:VerticalAlign.CENTER,
        margins:{top:0,bottom:0,left:0,right:0},
        children:[new Paragraph({spacing:{before:0,after:0,line:1},children:[new ImageRun({data:image,type:'png',transformation:{width:181,height:280}})]})]
      }))
    }
    rows.push(new TableRow({children:cells}))
  }

  const d=new Document({
    sections:[{
      properties:{
        page:{
          size:{width:11906,height:16838},
          margin:{top:283,bottom:283,left:283,right:283}
        }
      },
      children:[new Table({
        width:{size:100,type:WidthType.PERCENTAGE},
        borders:{
          top:{style:'nil',size:0,color:'FFFFFF'},
          bottom:{style:'nil',size:0,color:'FFFFFF'},
          left:{style:'nil',size:0,color:'FFFFFF'},
          right:{style:'nil',size:0,color:'FFFFFF'},
          insideHorizontal:{style:'nil',size:0,color:'FFFFFF'},
          insideVertical:{style:'nil',size:0,color:'FFFFFF'}
        },
        rows
      })]
    }]
  })
  const blob=await Packer.toBlob(d)
  const u=URL.createObjectURL(blob)
  const a=document.createElement('a')
  a.href=u
  a.download='crachas-cracha-fast.docx'
  a.click()
  setTimeout(()=>URL.revokeObjectURL(u),1000)
}