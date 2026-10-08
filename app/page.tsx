'use client';

import {useEffect,useMemo,useState} from 'react';
import {AlertTriangle,Boxes,Calculator,CheckCircle2,CircleDollarSign,ClipboardPaste,Download,Plus,RotateCcw,Settings2,Trash2,X} from 'lucide-react';
import * as XLSX from 'xlsx';

type C={id:string;date:string;no:string;tons:number;cu:number;au:number;ag:number;ni:number;sn:number;pd:number;cuPrice:number;auPrice:number;agPrice:number;pdPrice:number;note:string};
type R={cuHigh:number;cuLow:number;p1:number;p2:number;niGate:number;snGate:number;unitBase:number;specialBase:number;specialOn:boolean;specialCu:number;specialAg:number;auFees:{min:number;fee:number}[];auRates:{min:number;rate:number}[];agRates:{min:number;rate:number}[];pdFees:{min:number;fee:number}[]};
type X={tons:number;cu:number;au:number;ag:number;ni:number;sn:number;pd:number;effectiveCu:number;units:number;cuFee:number;penalty:number;auFee:number;auRate:number;agRate:number;pdFee:number;cuValue:number;auValue:number;agValue:number;pdValue:number;total:number};
type G={ids:string[];x:X;solo:number;gain:number;rate:number;reason:string};
type Goal={auMin:number;agMin:number;maxGroup:number};
type Disposition='single'|'hold';
type P={groups:G[];single:string[];hold:string[];base:number;total:number;gain:number;covered:number;coverage:number;goal:Goal;leftReasons:Record<string,string>};
type HalfState={mask:number;count:number;tons:number;auMass:number;agMass:number;auBalance:number;agBalance:number};
type GradeChoice={items:C[];count:number;tons:number;au:number;ag:number};

const R0:R={cuHigh:75,cuLow:70,p1:200,p2:500,niGate:3,snGate:1,unitBase:2000,specialBase:2500,specialOn:false,specialCu:30,specialAg:2000,auFees:[{min:200,fee:0},{min:150,fee:1},{min:100,fee:2},{min:70,fee:3},{min:50,fee:4},{min:40,fee:5},{min:30,fee:6},{min:20,fee:10},{min:10,fee:15},{min:5,fee:30}],auRates:[{min:100,rate:.96},{min:50,rate:.955},{min:30,rate:.95},{min:20,rate:.94},{min:5,rate:.935}],agRates:[{min:2000,rate:.92},{min:1500,rate:.91},{min:1300,rate:.90},{min:1000,rate:.89}],pdFees:[{min:30,fee:17},{min:20,fee:20},{min:10,fee:23},{min:5,fee:35},{min:3,fee:60}]};
const GOAL0:Goal={auMin:100,agMin:1000,maxGroup:30};
const demo:C[]=[['1','A01',24.8,68.4,92,1120,3.2,.7,4.2,78500,790,9.2,310,'Au接近100档'],['2','A02',25.2,80.6,115,820,1.1,1.3,6.1,78300,792,9.1,308,'高铜补位柜'],['3','A03',23.9,72.1,144,1680,2.9,.9,2.6,78600,788,9.3,312,'高金银柜'],['4','A04',24.4,76.8,82,960,4.4,1.2,11,78100,795,9.15,305,'待配柜'],['5','A05',25,69.2,158,1380,1.4,.5,3.1,78900,791,9.25,315,'高金易转手'],['6','A06',24.6,78.4,74,1096,2.1,.8,1.8,78400,789,9,309,'普通高铜柜']].map(a=>({id:a[0]as string,date:'2026-09-03',no:a[1]as string,tons:a[2]as number,cu:a[3]as number,au:a[4]as number,ag:a[5]as number,ni:a[6]as number,sn:a[7]as number,pd:a[8]as number,cuPrice:a[9]as number,auPrice:a[10]as number,agPrice:a[11]as number,pdPrice:a[12]as number,note:a[13]as string}));

const money=(v:number)=>new Intl.NumberFormat('zh-CN',{style:'currency',currency:'CNY',maximumFractionDigits:0}).format(v||0);
const num=(v:number,d=2)=>new Intl.NumberFormat('zh-CN',{maximumFractionDigits:d}).format(v||0);
const pct=(v:number,d=2)=>`${num(v,d)}%`;
const tier=(a:any[],v:number,k:string)=>a.find(x=>v>=x.min)?.[k]||0;
const auBandLabel=(v:number)=>v>=200?'Au ≥ 200':v>=150?'Au ≥ 150':v>=100?'Au ≥ 100':'未到Au 100档';
const agBandLabel=(v:number)=>v>=2000?'Ag ≥ 2000':v>=1500?'Ag ≥ 1500':v>=1300?'Ag ≥ 1300':v>=1000?'Ag ≥ 1000':'未到Ag 1000档';

function grades(a:C[]){
 const tons=a.reduce((s,c)=>s+c.tons,0);
 const w=(k:keyof C)=>tons?a.reduce((s,c)=>s+c.tons*Number(c[k]||0),0)/tons:0;
 return{tons,cu:w('cu'),au:w('au'),ag:w('ag'),ni:w('ni'),sn:w('sn'),pd:w('pd')};
}
function metrics(g:ReturnType<typeof grades>,r:R){
 const effectiveCu=g.cu+(g.ni>=r.niGate?g.ni:0)+(g.sn>=r.snGate?g.sn:0),units=g.au*10+Math.min(g.ag,1500),cuFee=Math.max(((r.specialOn&&g.cu<r.specialCu&&g.ag<r.specialAg)?r.specialBase:r.unitBase)-units,0),penalty=effectiveCu>=r.cuHigh?0:effectiveCu>=r.cuLow?(r.cuHigh-effectiveCu)*r.p1:(r.cuHigh-r.cuLow)*r.p1+(r.cuLow-effectiveCu)*r.p2;
 let agRate=0;if(g.ag>=1000)agRate=tier(r.agRates,g.ag,'rate');else if(g.ag>=100)agRate=.89-Math.ceil((1000-g.ag)/100)*.01;
 return{effectiveCu,units,cuFee,penalty,auFee:tier(r.auFees,g.au,'fee'),auRate:tier(r.auRates,g.au,'rate'),agRate,pdFee:tier(r.pdFees,g.pd,'fee')};
}
function calc(a:C[],r:R,unified=true):X{
 const g=grades(a),m=metrics(g,r);let cuValue=0,auValue=0,agValue=0,pdValue=0;
 for(const c of a){const z=unified?m:metrics(grades([c]),r);cuValue+=c.tons*c.cu/100*c.cuPrice-c.tons*(z.cuFee+z.penalty);auValue+=c.tons*c.au*Math.max(c.auPrice-z.auFee,0)*z.auRate;agValue+=c.tons*c.ag*c.agPrice*z.agRate;pdValue+=c.tons*c.pd*Math.max(c.pdPrice-z.pdFee,0)}
 return{...g,...m,cuValue,auValue,agValue,pdValue,total:cuValue+auValue+agValue+pdValue};
}
function group(a:C[],r:R,goal:Goal):G{
 const x=calc(a,r),solo=a.reduce((s,c)=>s+calc([c],r,false).total,0);
 return{ids:a.map(c=>c.id),x,solo,gain:x.total-solo,rate:solo?(x.total-solo)/solo*100:0,reason:`覆盖 ${a.length} 柜并同时满足 Au ≥ ${goal.auMin}、Ag ≥ ${goal.agMin}；同柜数时优先贴近门槛`};
}

function enumerateHalf(items:C[],goal:Goal){
 const size=2**items.length,all=Array.from({length:size}) as HalfState[],byCount:Array<HalfState[]> = Array.from({length:items.length+1},()=>[]);
 all[0]={mask:0,count:0,tons:0,auMass:0,agMass:0,auBalance:0,agBalance:0};byCount[0].push(all[0]);
 for(let mask=1;mask<size;mask++){
  const bit=mask&-mask,index=31-Math.clz32(bit),prev=all[mask^bit],c=items[index],tons=prev.tons+c.tons,auMass=prev.auMass+c.tons*c.au,agMass=prev.agMass+c.tons*c.ag;
  const state={mask,count:prev.count+1,tons,auMass,agMass,auBalance:auMass-goal.auMin*tons,agBalance:agMass-goal.agMin*tons};all[mask]=state;byCount[state.count].push(state);
 }
 return byCount;
}
function lastAtLeast(sorted:HalfState[],need:number){let lo=0,hi=sorted.length-1,ans=-1;while(lo<=hi){const mid=(lo+hi)>>1;if(sorted[mid].auBalance>=need){ans=mid;lo=mid+1}else hi=mid-1}return ans}
function choice(items:C[]):GradeChoice{const g=grades(items);return{items,count:items.length,tons:g.tons,au:g.au,ag:g.ag}}
function excess(a:GradeChoice,goal:Goal){const au=Math.max(0,a.au-goal.auMin)/Math.max(goal.auMin,1),ag=Math.max(0,a.ag-goal.agMin)/Math.max(goal.agMin,1);return{au,ag,total:au+ag,worst:Math.max(au,ag)}}
function betterChoice(a:GradeChoice,b:GradeChoice|null,goal:Goal){
 if(!b)return true;
 if(a.count!==b.count)return a.count>b.count;
 const ax=excess(a,goal),bx=excess(b,goal);
 if(Math.abs(ax.total-bx.total)>.000001)return ax.total<bx.total;
 if(Math.abs(ax.worst-bx.worst)>.000001)return ax.worst<bx.worst;
 if(Math.abs(ax.au-bx.au)>.000001)return ax.au<bx.au;
 if(Math.abs(ax.ag-bx.ag)>.000001)return ax.ag<bx.ag;
 if(Math.abs(a.tons-b.tons)>.000001)return a.tons>b.tons;
 return a.items.map(x=>x.no).join('|')<b.items.map(x=>x.no).join('|');
}
function meets(items:C[],goal:Goal){const g=grades(items);return items.length>=2&&g.au+1e-9>=goal.auMin&&g.ag+1e-9>=goal.agMin}
function improveChoice(start:GradeChoice,cs:C[],goal:Goal){
 let best=start;
 for(let pass=0;pass<3;pass++){
  let changed=false,inside=new Set(best.items.map(c=>c.id)),outside=cs.filter(c=>!inside.has(c.id));
  for(let i=0;i<best.items.length;i++)for(const add of outside){const next=[...best.items.slice(0,i),...best.items.slice(i+1),add];if(meets(next,goal)){const candidate=choice(next);if(betterChoice(candidate,best,goal)){best=candidate;changed=true}}}
  if(changed)continue;
  inside=new Set(best.items.map(c=>c.id));outside=cs.filter(c=>!inside.has(c.id));
  outer:for(let i=0;i<best.items.length;i++)for(let j=i+1;j<best.items.length;j++)for(let a=0;a<outside.length;a++)for(let b=a+1;b<outside.length;b++){
   const next=best.items.filter((_,k)=>k!==i&&k!==j).concat(outside[a],outside[b]);if(!meets(next,goal))continue;const candidate=choice(next);if(betterChoice(candidate,best,goal)){best=candidate;changed=true;break outer}
  }
  if(!changed)break;
 }
 return best;
}
function findFeasibleAtCount(cs:C[],goal:Goal,target:number){
 const n=cs.length;
 const split=Math.floor(n/2),leftItems=cs.slice(0,split),rightItems=cs.slice(split),left=enumerateHalf(leftItems,goal),rightRaw=enumerateHalf(rightItems,goal);
 const right=rightRaw.map(states=>{
  const sorted=[...states].sort((a,b)=>b.auBalance-a.auBalance||b.agBalance-a.agBalance),prefixBest:number[]=[];let best=-Infinity,bestIndex=-1;
  sorted.forEach((s,i)=>{if(s.agBalance>best){best=s.agBalance;bestIndex=i}prefixBest[i]=bestIndex});return{sorted,prefixBest};
 });
 let best:GradeChoice|null=null;
 for(let lc=Math.max(0,target-rightItems.length);lc<=Math.min(leftItems.length,target);lc++){
  const rc=target-lc,index=right[rc];if(!index||!index.sorted.length)continue;
  for(const l of left[lc]){
   const last=lastAtLeast(index.sorted,-l.auBalance);if(last<0)continue;const ri=index.prefixBest[last],rr=index.sorted[ri];if(rr.agBalance+l.agBalance<-1e-9)continue;
   const items:C[]=[];for(let i=0;i<leftItems.length;i++)if(l.mask&2**i)items.push(leftItems[i]);for(let i=0;i<rightItems.length;i++)if(rr.mask&2**i)items.push(rightItems[i]);
   const candidate=choice(items);if(betterChoice(candidate,best,goal))best=candidate;
  }
 }
 return best;
}
function findBestGradeGroup(cs:C[],rawGoal:Goal){
 const goal={...rawGoal,maxGroup:Math.min(30,Math.max(2,Math.floor(rawGoal.maxGroup||30)))},n=cs.length;
 if(n<2)return[] as C[];
 for(let target=Math.min(goal.maxGroup,n);target>=2;target--){
  const best=findFeasibleAtCount(cs,goal,target);if(!best)continue;
  return improveChoice(best,cs,goal).items;
 }
 return[] as C[];
}
function partitionQuality(parts:C[][],goal:Goal){
 const gaps=parts.map(items=>excess(choice(items),goal)),count=parts.reduce((s,x)=>s+x.length,0)||1;
 return{worst:Math.max(0,...gaps.map(x=>x.worst)),average:gaps.reduce((s,x,i)=>s+x.total*parts[i].length,0)/count,groups:parts.length};
}
function betterPartition(a:ReturnType<typeof partitionQuality>,b:ReturnType<typeof partitionQuality>){
 if(Math.abs(a.worst-b.worst)>.000001)return a.worst<b.worst;
 if(Math.abs(a.average-b.average)>.000001)return a.average<b.average;
 return a.groups>b.groups;
}
function bestSplit(items:C[],goal:Goal){
 if(items.length<4)return null;
 const total=grades(items),totalAu=total.tons*(total.au-goal.auMin),totalAg=total.tons*(total.ag-goal.agMin);let best:C[][]|null=null,bestScore:ReturnType<typeof partitionQuality>|null=null;
 const consider=(left:C[])=>{if(left.length<2||left.length>items.length-2)return;const ids=new Set(left.map(c=>c.id)),right=items.filter(c=>!ids.has(c.id));if(!meets(left,goal)||!meets(right,goal))return;const parts=[left,right],score=partitionQuality(parts,goal);if(!bestScore||betterPartition(score,bestScore)){best=parts;bestScore=score}};
 const maxSmall=Math.min(4,Math.floor(items.length/2));
 for(let size=2;size<=maxSmall;size++){
  const picked:number[]=[];
  const walk=(start:number,need:number,tons:number,auMass:number,agMass:number)=>{
   if(!need){const auBalance=auMass-goal.auMin*tons,agBalance=agMass-goal.agMin*tons;if(auBalance>=-1e-9&&agBalance>=-1e-9&&auBalance<=totalAu+1e-9&&agBalance<=totalAg+1e-9)consider(picked.map(i=>items[i]));return}
   for(let i=start;i<=items.length-need;i++){const c=items[i];picked.push(i);walk(i+1,need-1,tons+c.tons,auMass+c.tons*c.au,agMass+c.tons*c.ag);picked.pop()}
  };
  walk(0,size,0,0,0);
 }
 if(items.length>8){const strategic=[Math.floor(items.length/3),Math.floor(items.length/2),Math.ceil(items.length/2)];for(const size of [...new Set(strategic.filter(x=>x>=2&&x<=items.length-2))]){const candidate=findFeasibleAtCount(items,goal,size);if(candidate)consider(candidate.items)}}
 return best;
}
function partitionSelected(selected:C[],goal:Goal){
 let parts=selected.length?[selected]:[] as C[][];
 for(let pass=0;pass<15;pass++){
  let next:C[][]|null=null,nextScore:ReturnType<typeof partitionQuality>|null=null;
  for(let i=0;i<parts.length;i++){const split=bestSplit(parts[i],goal);if(!split)continue;const candidate=[...parts.slice(0,i),...split,...parts.slice(i+1)],score=partitionQuality(candidate,goal);if(!nextScore||betterPartition(score,nextScore)){next=candidate;nextScore=score}}
  if(!next)break;parts=next;
 }
 return parts.sort((a,b)=>b.length-a.length||grades(a).au-grades(b).au);
}
function optimize(cs:C[],r:R,goal:Goal,disposition:Disposition):P{
 const selected=findBestGradeGroup(cs,goal),parts=partitionSelected(selected,goal),groups=parts.map(items=>group(items,r,goal)),used=new Set(selected.map(c=>c.id)),left=cs.filter(c=>!used.has(c.id)),single=disposition==='single'?left.map(c=>c.id):[],hold=disposition==='hold'?left.map(c=>c.id):[];
 const included=[...selected,...left.filter(c=>single.includes(c.id))],base=included.reduce((s,c)=>s+calc([c],r,false).total,0),total=groups.reduce((s,g)=>s+g.x.total,0)+left.filter(c=>single.includes(c.id)).reduce((s,c)=>s+calc([c],r,false).total,0),leftReasons:Record<string,string>={};
 for(const c of left){if(selected.length>=goal.maxGroup){leftReasons[c.id]=`本批已达到入组上限 ${goal.maxGroup} 柜`;continue}const projected=grades([...selected,c]),reasons=[];if(projected.au<goal.auMin)reasons.push(`并入全部组合后 Au 仅 ${num(projected.au,2)}`);if(projected.ag<goal.agMin)reasons.push(`并入全部组合后 Ag 仅 ${num(projected.ag,2)}`);leftReasons[c.id]=reasons.length?reasons.join('、'):'全局比较后未进入最大覆盖方案'}
 return{groups,single,hold,base,total,gain:total-base,covered:selected.length,coverage:cs.length?selected.length/cs.length*100:0,goal,leftReasons};
}

function Field({label,value,onChange,min,max}:{label:string,value:any,onChange:(v:string)=>void;min?:number;max?:number}){return <label className="field"><span>{label}</span><input value={value} type="number" step="any" min={min} max={max} onChange={e=>onChange(e.target.value)}/></label>}

export default function App(){
 const[cs,setCs]=useState<C[]>(demo),[r,setR]=useState<R>(R0),[goal,setGoal]=useState<Goal>(GOAL0),[goalDraft,setGoalDraft]=useState<Goal>(GOAL0),[disposition,setDisposition]=useState<Disposition>('single'),[tab,setTab]=useState('plan'),[showRules,setShowRules]=useState(false),[manual,setManual]=useState<Record<string,string>>({});
 const plan=useMemo(()=>optimize(cs,r,goal,disposition),[cs,r,goal,disposition]);
 useEffect(()=>{try{const s=JSON.parse(localStorage.getItem('waste-copper-state')||'null');if(s){if(Array.isArray(s.cs))setCs(s.cs);if(s.r)setR(s.rulesVersion>=2?s.r:{...s.r,auRates:R0.auRates});const loaded={...GOAL0,...s.goal};setGoal(loaded);setGoalDraft(loaded);if(s.disposition)setDisposition(s.disposition)}}catch{}},[]);
 useEffect(()=>localStorage.setItem('waste-copper-state',JSON.stringify({cs,r,goal,disposition,rulesVersion:3})),[cs,r,goal,disposition]);
 function applyGoal(){const next={auMin:Math.max(0,Number(goalDraft.auMin)||100),agMin:Math.max(0,Number(goalDraft.agMin)||1000),maxGroup:Math.min(30,Math.max(2,Math.floor(Number(goalDraft.maxGroup)||30)))};setGoalDraft(next);setGoal(next)}
 function exportExcel(){
  const wb=XLSX.utils.book_new(),add=(n:string,rows:any[])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),n),ex=(x:X)=>({吨数:x.tons,Cu:x.cu,Au:x.au,Ag:x.ag,Ni:x.ni,Sn:x.sn,Pd:x.pd,有效铜:x.effectiveCu,单位数:x.units,铜加工费:x.cuFee,铜扣款:x.penalty,金加工费:x.auFee,金系数:x.auRate,银系数:x.agRate,钯加工费:x.pdFee,铜价值:x.cuValue,金价值:x.auValue,银价值:x.agValue,钯价值:x.pdValue,总价值:x.total}),selectedIds=new Set(plan.groups.flatMap(g=>g.ids)),selected=cs.filter(c=>selectedIds.has(c.id)),selectedGrade=grades(selected);
  add('批次信息',[{优化目标:'金银达标后覆盖柜数最多，并拆成尽可能多的独立达标组',黄金最低品位:goal.auMin,白银最低品位:goal.agMin,本批最多入组柜数:goal.maxGroup,每组最少柜数:2,组合数量限制:'不限',货柜总数:cs.length,入组柜数:plan.covered,覆盖率:plan.coverage,剩余柜处理:disposition==='single'?'单独结算':'留待下批',生成时间:new Date().toLocaleString()}]);
  add('货柜原始数据与点价',cs.map(({id:_,...c})=>c));add('单柜结算',cs.map(c=>({柜号:c.no,...ex(calc([c],r,false))})));
  add('推荐组合',plan.groups.map((g,i)=>({组合:`组合${i+1}`,柜号:g.ids.map(id=>cs.find(c=>c.id===id)?.no).join('+'),入组柜数:g.ids.length,Au目标:goal.auMin,Ag目标:goal.agMin,Au超出:g.x.au-goal.auMin,Ag超出:g.x.ag-goal.agMin,黄金档位:auBandLabel(g.x.au),白银档位:agBandLabel(g.x.ag),...ex(g.x),说明:g.reason})));
  add('组合加权过程',plan.groups.flatMap((g,i)=>g.ids.map(id=>{const c=cs.find(x=>x.id===id)!;return{组合:`组合${i+1}`,柜号:c.no,吨数:c.tons,Au:c.au,Au加权量:c.tons*c.au,Ag:c.ag,Ag加权量:c.tons*c.ag,是否入组:'是'}})));
  add('组合前后结算对比',plan.groups.flatMap((g,i)=>g.ids.map(id=>{const c=cs.find(x=>x.id===id)!,solo=calc([c],r,false),combo=allocatedValue(c,g.x);return{组合:`组合${i+1}`,柜号:c.no,单独价值:solo.total,组合分摊价值:combo,差额:combo-solo.total}})));
  add('未入组清单',[...plan.single,...plan.hold].map(id=>{const c=cs.find(x=>x.id===id)!;const projected=grades([...selected,c]);return{柜号:c.no,吨数:c.tons,Au:c.au,Ag:c.ag,加入后Au:projected.au,加入后Ag:projected.ag,处理:plan.single.includes(id)?'单独结算':'留待下批',原因:plan.leftReasons[id]}}));
  add('组合汇总',[{组合数:plan.groups.length,入组柜数:plan.covered,总柜数:cs.length,覆盖率:plan.coverage,全部入组柜加权Au:selectedGrade.au,全部入组柜加权Ag:selectedGrade.ag,单独结算基准:plan.base,组合后结算:plan.total,金额差异:plan.gain,备注:'金额仅供结算核对，不参与推荐组合选择'}]);
  const manualRows=cs.map(c=>{const gi=plan.groups.findIndex(g=>g.ids.includes(c.id));return{柜号:c.no,人工安排:manual[c.id]||(gi>=0?`组合${gi+1}`:plan.hold.includes(c.id)?'留待下批':'单独结算')}});add('人工调整',manualRows);add('计价与优化规则',[{优化规则JSON:JSON.stringify(goal,null,2),计价规则JSON:JSON.stringify(r,null,2)}]);
  XLSX.writeFile(wb,`废铜金银达标组合_${new Date().toISOString().slice(0,10)}.xlsx`);
 }
 return <main><header className="topbar"><div className="brand"><div className="brandmark"><Boxes size={20}/></div><div><h1>废铜组合销售与结算</h1><p>金银达标后，优先让更多柜进入组合</p></div></div><div className="actions"><button className="ghost" onClick={()=>setShowRules(!showRules)}><Settings2 size={16}/>计价规则</button><button className="primary" onClick={exportExcel}><Download size={16}/>导出 Excel</button></div></header><section className="workspace"><aside className="sidebar"><nav>{[['plan','达标方案',Calculator],['data','货柜与点价',Boxes],['compare','结算核对',CircleDollarSign]].map(([k,t,I]:any)=><button key={k} className={tab===k?'active':''} onClick={()=>setTab(k)}><I size={17}/>{t}</button>)}</nav><div className="local"><b>本地保存</b><span>数据仅保存在当前设备</span></div></aside><div className="content"><section className="controls"><div><span className="eyebrow">优化目标</span><div className="target-badge"><CheckCircle2 size={15}/>金银达标 · 多组不限</div></div><Field label="黄金最低品位 Au g/t" value={goalDraft.auMin} min={0} onChange={v=>setGoalDraft(x=>({...x,auMin:+v}))}/><Field label="白银最低品位 Ag g/t" value={goalDraft.agMin} min={0} onChange={v=>setGoalDraft(x=>({...x,agMin:+v}))}/><Field label="本批最多入组柜数" value={goalDraft.maxGroup} min={2} max={30} onChange={v=>setGoalDraft(x=>({...x,maxGroup:+v}))}/><label className="field"><span>剩余柜处理</span><select value={disposition} onChange={e=>setDisposition(e.target.value as Disposition)}><option value="single">单独结算</option><option value="hold">留待下批</option></select></label><button className="primary update-plan" onClick={applyGoal}><Calculator size={16}/>更新方案</button></section><p className="objective-note"><AlertTriangle size={14}/>每组至少2柜、组数不限；点价和人民币金额只用于结算核对，不参与推荐方案选择。</p>{showRules&&<RulesBuffered r={r} setR={setR}/>} {tab==='plan'&&<PlanView plan={plan} cs={cs}/>} {tab==='data'&&<DataBuffered cs={cs} setCs={setCs}/>} {tab==='compare'&&<Compare plan={plan} cs={cs} r={r} manual={manual} setManual={setManual}/>}</div></section></main>;
}

function allocatedValue(c:C,m:X){return c.tons*c.cu/100*c.cuPrice-c.tons*(m.cuFee+m.penalty)+c.tons*c.au*Math.max(c.auPrice-m.auFee,0)*m.auRate+c.tons*c.ag*c.agPrice*m.agRate+c.tons*c.pd*Math.max(c.pdPrice-m.pdFee,0)}
function Title({over,children,right}:{over:string;children:any;right?:any}){return <div className="title"><div><span className="eyebrow">{over}</span><h2>{children}</h2></div>{right}</div>}
function Kpi({label,value,sub,good=false}:any){return <article className={`kpi ${good?'good':''}`}><span>{label}</span><strong>{value}</strong><small>{sub}</small></article>}
function PlanView({plan,cs}:{plan:P;cs:C[]}){
 const all=plan.groups.flatMap(g=>g.ids),selected=cs.filter(c=>all.includes(c.id)),overall=grades(selected),quality=plan.groups.length?partitionQuality(plan.groups.map(g=>cs.filter(c=>g.ids.includes(c.id))),plan.goal):null;
 return <><Title over="本批推荐" right={<span className="status"><CheckCircle2 size={16}/>已全局计算 {cs.length} 柜</span>}>金银达标多组合方案</Title><section className="kpis"><Kpi good label="达标入组" value={`${plan.covered} / ${cs.length} 柜`} sub={`覆盖率 ${pct(plan.coverage,0)}`}/><Kpi label="达标组合" value={`${plan.groups.length} 组`} sub="每组至少2柜，组数不限"/><Kpi label="全部入组加权品位" value={plan.covered?`Au ${num(overall.au,2)}`:'—'} sub={plan.covered?`Ag ${num(overall.ag,2)} g/t`:`目标 Au ≥ ${plan.goal.auMin} · Ag ≥ ${plan.goal.agMin}`}/><Kpi label="最远门槛偏差" value={quality?pct(quality.worst*100,2):'—'} sub="先覆盖更多柜，再拆分达标组并贴近门槛"/></section><div className="cards">{plan.groups.length?plan.groups.map((g,i)=><article className="group" key={i}><div className="grouphead"><div><small>组合 {String(i+1).padStart(2,'0')} · 达标贴线</small><h3>{g.ids.map(id=>cs.find(c=>c.id===id)?.no).join(' + ')}</h3></div><b>{g.ids.length} 柜</b></div><div className="grades"><span><b>{num(g.x.au,2)}</b>Au g/t</span><span><b>{num(g.x.ag,2)}</b>Ag g/t</span><span><b>{num(g.x.tons,2)}</b>总吨数</span><span><b>{num(g.x.units,0)}</b>单位数</span></div><p className="reason"><CheckCircle2 size={14}/>{g.reason}</p><footer><span>组合结算 {money(g.x.total)}</span><span>单柜基准 {money(g.solo)}</span><b>金额仅供参考 {g.gain>=0?'+':''}{money(g.gain)}</b></footer></article>):<div className="empty"><AlertTriangle/><h3>暂时没有达标组合</h3><p>每组至少需要2柜，并同时达到 Au ≥ {plan.goal.auMin}、Ag ≥ {plan.goal.agMin}。</p></div>}<LeftList title={plan.hold.length?'留待下批':'单独结算'} ids={plan.hold.length?plan.hold:plan.single} cs={cs} plan={plan}/></div></>;
}
function LeftList({title,ids,cs,plan}:{title:string;ids:string[];cs:C[];plan:P}){return <article className="list"><div className="listhead"><h3>{title}</h3><span>{ids.length} 柜</span></div>{ids.length?ids.map(id=>{const c=cs.find(x=>x.id===id)!;return <div className="listrow left-reason" key={id}><div><b>{c.no}</b><small>Au {num(c.au,2)} · Ag {num(c.ag,2)}</small></div><span>{plan.leftReasons[id]}</span></div>}):<p className="muted">全部货柜已进入达标组合</p>}</article>}

function DataBuffered({cs,setCs}:{cs:C[];setCs:(v:any)=>void}){const[draft,setDraft]=useState<C[]>(cs),[saved,setSaved]=useState(false);useEffect(()=>setDraft(cs),[cs]);const dirty=JSON.stringify(draft)!==JSON.stringify(cs),upd=(id:string,k:keyof C,v:string)=>{setSaved(false);setDraft(a=>a.map(c=>c.id===id?{...c,[k]:['date','no','note'].includes(k)?v:Number(v)}:c))},add=()=>{setSaved(false);setDraft(a=>a.length>=30?a:[...a,{...demo[0],id:crypto.randomUUID(),no:`A${String(a.length+1).padStart(2,'0')}`,note:''}])},del=(id:string)=>{setSaved(false);setDraft(a=>a.filter(c=>c.id!==id))},reset=()=>{if(confirm('恢复默认示例数据？')){setDraft(demo);setSaved(false)}};return <><Data cs={draft} setCs={setDraft} upd={upd} add={add} del={del} reset={reset}/><div className={`apply-bar ${dirty?'dirty':''}`}><div><b>{dirty?'有尚未应用的数据修改':'货柜数据已应用'}</b><span>{dirty?'输入期间不会触发组合计算，避免页面卡顿。':'方案已使用当前货柜数据计算。'}</span></div><button className="primary" disabled={!dirty} onClick={()=>{setCs(draft);setSaved(true)}}><Calculator size={16}/>{saved?'已更新':'应用数据并更新方案'}</button></div></>}
function Data({cs,setCs,upd,add,del,reset}:any){
 const[paste,setPaste]=useState(''),[replace,setReplace]=useState(false),[message,setMessage]=useState('');const cols:[keyof C,string][]=[['no','柜号'],['tons','吨数'],['cu','Cu%'],['au','Au g/t'],['ag','Ag g/t'],['ni','Ni%'],['sn','Sn%'],['pd','Pd g/t'],['cuPrice','铜价 元/t'],['auPrice','金价 元/g'],['agPrice','银价 元/g'],['pdPrice','钯价 元/g']];
 const aliases:Record<string,keyof C>={'日期':'date','date':'date','柜号':'no','货柜号':'no','no':'no','吨数':'tons','重量':'tons','tons':'tons','cu':'cu','cu%':'cu','铜':'cu','au':'au','au g/t':'au','金':'au','ag':'ag','ag g/t':'ag','银':'ag','ni':'ni','ni%':'ni','镍':'ni','sn':'sn','sn%':'sn','锡':'sn','pd':'pd','pd g/t':'pd','钯':'pd','铜价':'cuPrice','铜价 元/t':'cuPrice','cuprice':'cuPrice','金价':'auPrice','金价 元/g':'auPrice','auprice':'auPrice','银价':'agPrice','银价 元/g':'agPrice','agprice':'agPrice','钯价':'pdPrice','钯价 元/g':'pdPrice','pdprice':'pdPrice','备注':'note','note':'note'};
 function importPaste(){setMessage('');const rows=paste.trim().split(/\r?\n/).map(r=>r.split('\t').map(v=>v.trim())).filter(r=>r.some(Boolean));if(!rows.length){setMessage('请先从 Excel 复制数据并粘贴到上方。');return}const normalized=rows[0].map(v=>v.toLowerCase().replace(/[（(].*?[）)]/g,'').trim()),hasHeader=normalized.some(v=>aliases[v]);const fixed:(keyof C)[]=['date','no','tons','cu','au','ag','ni','sn','pd','cuPrice','auPrice','agPrice','pdPrice','note'];const keys=hasHeader?normalized.map(v=>aliases[v]||null):fixed,data=rows.slice(hasHeader?1:0),defaults=cs[0]||demo[0],imported:C[]=[],errors:string[]=[];data.forEach((row,idx)=>{const raw:any={};keys.forEach((k,i)=>{if(k)raw[k]=row[i]??''});if(!raw.no||!raw.tons){errors.push(`第 ${idx+(hasHeader?2:1)} 行缺少柜号或吨数`);return}const n=(k:keyof C,f=0)=>{const v=String(raw[k]??'').replace(/,/g,'').replace('%','');return v===''?f:Number(v)};const c:C={id:crypto.randomUUID(),date:raw.date||new Date().toISOString().slice(0,10),no:String(raw.no),tons:n('tons'),cu:n('cu'),au:n('au'),ag:n('ag'),ni:n('ni'),sn:n('sn'),pd:n('pd'),cuPrice:n('cuPrice',defaults.cuPrice),auPrice:n('auPrice',defaults.auPrice),agPrice:n('agPrice',defaults.agPrice),pdPrice:n('pdPrice',defaults.pdPrice),note:String(raw.note||'')};if(!Number.isFinite(c.tons)||c.tons<=0||[c.cu,c.au,c.ag,c.ni,c.sn,c.pd,c.cuPrice,c.auPrice,c.agPrice,c.pdPrice].some(v=>!Number.isFinite(v))){errors.push(`第 ${idx+(hasHeader?2:1)} 行含无效数字`);return}imported.push(c)});if((replace?0:cs.length)+imported.length>30){setMessage(`导入后将超过 30 柜，请减少 ${(replace?0:cs.length)+imported.length-30} 行。`);return}if(!imported.length){setMessage(errors.join('；')||'未识别到有效数据。');return}const duplicate=[...(replace?[]:cs.map((c:C)=>c.no)),...imported.map(c=>c.no)].filter((v,i,a)=>a.indexOf(v)!==i);if(duplicate.length){setMessage(`柜号重复：${[...new Set(duplicate)].join('、')}`);return}setPaste('');setMessage(`已导入 ${imported.length} 柜${errors.length?`；另有 ${errors.length} 行未导入`:''}`);(window as any).__setImported?.(imported,replace)}
 useEffect(()=>{(window as any).__setImported=(items:C[],isReplace:boolean)=>setCs((old:C[])=>isReplace?items:[...old,...items]);return()=>{delete (window as any).__setImported}},[setCs]);
 return <><Title over="原始数据" right={<div className="actions"><button className="ghost" onClick={reset}><RotateCcw size={15}/>恢复示例</button><button className="primary" onClick={add}><Plus size={16}/>逐柜新增</button></div>}>货柜品位与每柜点价</Title><section className="import-card"><div className="import-head"><div className="import-icon"><ClipboardPaste size={19}/></div><div><h3>从 Excel 批量粘贴</h3><p>直接复制 Excel 中的多行数据，保留表头效果最佳</p></div><label className="replace-check"><input type="checkbox" checked={replace} onChange={e=>setReplace(e.target.checked)}/>替换现有货柜</label></div><textarea value={paste} onChange={e=>setPaste(e.target.value)} placeholder={'日期\t柜号\t吨数\tCu%\tAu g/t\tAg g/t\tNi%\tSn%\tPd g/t\t铜价 元/t\t金价 元/g\t银价 元/g\t钯价 元/g\t备注\n2026-09-03\tA07\t24.8\t72.5\t108\t1280\t3.1\t0.8\t4.5\t78500\t790\t9.2\t310\t待配柜'} /><div className="import-foot"><span>必填：柜号、吨数。未填写点价时沿用首柜价格。</span><div>{paste&&<button className="clear-paste" onClick={()=>{setPaste('');setMessage('')}}><X size={14}/>清空</button>}<button className="primary" onClick={importPaste}><ClipboardPaste size={15}/>识别并导入</button></div></div>{message&&<p className={message.startsWith('已导入')?'import-ok':'import-error'}>{message}</p>}</section><div className="entry-label"><span>方式二 · 逐柜输入</span><small>可直接在表格单元格内修改，最多30柜</small></div><div className="table"><table><thead><tr>{cols.map(c=><th key={c[0]}>{c[1]}</th>)}<th>操作</th></tr></thead><tbody>{cs.map((c:C)=><tr key={c.id}>{cols.map(([k])=><td key={k}><input value={c[k]} onChange={e=>upd(c.id,k,e.target.value)}/></td>)}<td><button className="icon" onClick={()=>del(c.id)}><Trash2 size={15}/></button></td></tr>)}</tbody></table></div><p className="hint"><AlertTriangle size={14}/>每柜价格独立保存；点价只影响结算金额，不影响入组推荐。</p></>;
}

function Compare({plan,cs,r,manual,setManual}:{plan:P;cs:C[];r:R;manual:Record<string,string>;setManual:any}){
 const auto=(c:C)=>{const gi=plan.groups.findIndex(g=>g.ids.includes(c.id));return gi>=0?`组合${gi+1}`:plan.hold.includes(c.id)?'留待下批':'单独结算'},status=(c:C)=>manual[c.id]||auto(c),manualLabels=Array.from({length:Math.max(plan.groups.length,Math.floor(cs.length/2),1)},(_,i)=>`组合${i+1}`),entries=manualLabels.map(label=>({label,items:cs.filter(c=>status(c)===label)})).filter(x=>x.items.length),invalid=entries.filter(x=>!meets(x.items,plan.goal)),valid=entries.length>0&&!invalid.length,manualMetrics=new Map(entries.map(x=>[x.label,meets(x.items,plan.goal)?calc(x.items,r):null]));
 return <><Title over="人工调整与结算核对" right={<button className="ghost" onClick={()=>setManual({})}><RotateCcw size={15}/>恢复系统推荐</button>}>单柜与各组合逐项对比</Title><div className={`manual-check ${valid?'valid':'invalid'}`}><div>{valid?<CheckCircle2/>:<AlertTriangle/>}<span><b>{valid?'全部人工组合达标':'人工方案存在未达标组合'}</b><small>{entries.length}组 · {entries.reduce((s,x)=>s+x.items.length,0)}柜进入人工组合</small></span></div><p>{valid?'每组均不少于2柜，且Au、Ag均未跌档。':invalid.map(x=>`${x.label}（${x.items.length}柜，Au ${num(grades(x.items).au,2)}，Ag ${num(grades(x.items).ag,2)}）`).join('；')||'请至少建立一个组合。'}</p></div><div className="summary"><span>单独结算基准<b>{money(plan.base)}</b></span><i>→</i><span>系统组合结算<b>{money(plan.total)}</b></span><strong>金额仅供核对 {plan.gain>=0?'+':''}{money(plan.gain)}</strong></div><div className="table"><table><thead><tr><th>柜号</th><th>人工安排</th><th>单柜系数 / 费用</th><th>人工组合系数 / 费用</th><th>单独价值</th><th>安排后价值</th><th>差额</th></tr></thead><tbody>{cs.map(c=>{const s=calc([c],r,false),st=status(c),m=manualMetrics.get(st)||null,cv=m?allocatedValue(c,m):st==='留待下批'?null:st.startsWith('组合')?null:s.total;return <tr key={c.id}><td><b>{c.no}</b></td><td><select value={st} onChange={e=>setManual((x:any)=>({...x,[c.id]:e.target.value}))}>{manualLabels.map(label=><option key={label}>{label}</option>)}<option>单独结算</option><option>留待下批</option></select></td><td className="rulecell">Au {(s.auRate*100).toFixed(1)}% · Ag {(s.agRate*100).toFixed(0)}%<br/>Cu {num(s.cuFee,0)} / Au {s.auFee} / Pd {s.pdFee}</td><td className="rulecell">{m?<>Au {(m.auRate*100).toFixed(1)}% · Ag {(m.agRate*100).toFixed(0)}%<br/>Cu {num(m.cuFee,0)} / Au {m.auFee} / Pd {m.pdFee}</>:'—'}</td><td>{money(s.total)}</td><td>{cv===null?'—':money(cv)}</td><td className={cv!==null&&cv-s.total>=0?'pos':'neg'}>{cv===null?'—':`${cv-s.total>=0?'+':''}${money(cv-s.total)}`}</td></tr>})}</tbody></table></div><p className="hint">人工安排可以使用多个组合；每组至少2柜并同时达到Au、Ag门槛。</p></>;
}

function RulesBuffered({r,setR}:{r:R;setR:(r:R)=>void}){const[draft,setDraft]=useState(r);useEffect(()=>setDraft(r),[r]);const dirty=JSON.stringify(draft)!==JSON.stringify(r);return <div><Rules r={draft} setR={setDraft}/><div className={`rules-apply ${dirty?'dirty':''}`}><span>{dirty?'规则已修改，应用后更新结算金额':'当前计价规则已应用'}</span><button className="primary" disabled={!dirty} onClick={()=>setR(draft)}><Calculator size={15}/>应用规则并更新</button></div></div>}
function Rules({r,setR}:{r:R;setR:(r:R)=>void}){const f=(k:keyof R,v:string)=>setR({...r,[k]:Number(v)});return <section className="rules"><div className="ruleshead"><div><span className="eyebrow">集中参数中心</span><h3>计价规则</h3></div><span>计价规则只用于结算，不参与柜数推荐</span></div><div className="rulegrid"><Field label="Cu免扣线 %" value={r.cuHigh} onChange={v=>f('cuHigh',v)}/><Field label="Cu重扣线 %" value={r.cuLow} onChange={v=>f('cuLow',v)}/><Field label="70–75扣/点" value={r.p1} onChange={v=>f('p1',v)}/><Field label="70以下扣/点" value={r.p2} onChange={v=>f('p2',v)}/><Field label="Ni门槛 %" value={r.niGate} onChange={v=>f('niGate',v)}/><Field label="Sn门槛 %" value={r.snGate} onChange={v=>f('snGate',v)}/><Field label="加工费基数" value={r.unitBase} onChange={v=>f('unitBase',v)}/></div><details><summary>编辑金银钯阶梯参数</summary><textarea value={JSON.stringify({auFees:r.auFees,auRates:r.auRates,agRates:r.agRates,pdFees:r.pdFees},null,2)} onChange={e=>{try{setR({...r,...JSON.parse(e.target.value)})}catch{}}}/><p>黄金系数：5–&lt;20 为93.5%，20–&lt;30为94%，30–&lt;50为95%，50–&lt;100为95.5%，100以上为96%。</p></details></section>}
