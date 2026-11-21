export const pods=[{id:'orange',name:'Team 1',color:'#FFC000',anchor:'2026-11-03'},{id:'green',name:'Team 2',color:'#92D050',anchor:'2026-11-01'},{id:'blue',name:'Team 3',color:'#00B0F0',anchor:'2026-11-05'},{id:'red',name:'Team 4',color:'#FF0000',anchor:'2026-10-30'}];
// Real staff identities are loaded from the authenticated roster API at runtime.
export const staff=[];
const day=(s)=>Math.floor((Date.parse(s+'T00:00:00Z')-Date.UTC(1970,0,1))/86400000);
export function cycle(date,anchor){let phase=((day(date)-day(anchor))%8+8)%8;return phase<2?'M':phase<4?'N':'OFF'}
function memberForDate(state,id,date){let person=staff.find(x=>x.id===id),month=date.slice(0,7),current=state.staffMonthly?.[`${month}:${id}`];if(current)return {...(person||{id}),...current};if(!person)return null;let prior=Object.entries(state.staffMonthly||{}).filter(([entry,value])=>entry.endsWith(`:${id}`)&&entry.slice(0,7)<month&&value?.podId).sort(([a],[b])=>a.localeCompare(b)).at(-1)?.[1];return prior?{...person,podId:prior.podId,rowOrder:prior.rowOrder}:person}
export function cellFor(state,staffId,date){let key=`${date}:${staffId}`;return state.overrides[key]||state.assignments[key]||(()=>{let member=memberForDate(state,staffId,date),p=pods.find(x=>x.id===member?.podId);if(!member||!p||date<(member.effectiveFrom||'0000-00-00'))return 'OFF';return cycle(date,p.anchor)})()}
export function setCell(state,id,date,value){let key=`${date}:${id}`;if(!['M','N','OFF','LEAVE'].includes(value))return;let member=memberForDate(state,id,date),p=pods.find(x=>x.id===member?.podId),baseline=state.assignments[key]||(p&&date>=(member?.effectiveFrom||'0000-00-00')?cycle(date,p.anchor):'OFF');if(value===baseline)delete state.overrides[key];else state.overrides[key]=value;}
export function initialState(){return {assignments:{},overrides:{},staffMonthly:{},coverAssignments:{},coverTombstones:[]}}
export function roleFor(s,month){return s.roleByMonth?.[month]||s.role}
export function nameFor(s){return s.shortName||s.name}
export function buildMonth(state,month){return state}
export function generateMonth(state,previous,next){state.staffMonthly||={};let [y,m]=next.split('-').map(Number),n=new Date(y,m,0).getDate();for(let s of staff){if(state.inactiveStaffIds?.includes(s.id))continue;let prior=state.staffMonthly[`${previous}:${s.id}`];if(prior)state.staffMonthly[`${next}:${s.id}`]={...prior};let member=state.staffMonthly[`${next}:${s.id}`]||s,p=pods.find(x=>x.id===member.podId);if(!p)continue;for(let d=1;d<=n;d++){let date=`${next}-${String(d).padStart(2,'0')}`;state.assignments[`${date}:${s.id}`]=cycle(date,p.anchor)}}return state}
export function monthName(m){return new Date(m+'-01T00:00:00Z').toLocaleString('en',{month:'long',timeZone:'UTC'})}
