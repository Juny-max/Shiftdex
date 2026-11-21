export const pods=[{id:'orange',name:'Team 1',color:'#FFC000',anchor:'2026-11-03'},{id:'green',name:'Team 2',color:'#92D050',anchor:'2026-11-01'},{id:'blue',name:'Team 3',color:'#00B0F0',anchor:'2026-11-05'},{id:'red',name:'Team 4',color:'#FF0000',anchor:'2026-10-30'}];
// Names and shift assignments must come from the authenticated API, never the public portal bundle.
export const staff=[];
const day=s=>Math.floor(Date.parse(s+'T00:00:00Z')/86400000);
export function cycle(date,anchor){let phase=((day(date)-day(anchor))%8+8)%8;return phase<2?'M':phase<4?'N':'OFF'}
function memberForDate(state,id,date){let person=staff.find(x=>x.id===id);if(!person)return null;let month=date.slice(0,7),current=state.staffMonthly?.[`${month}:${id}`];if(current)return {...person,...current};let prior=Object.entries(state.staffMonthly||{}).filter(([entry,value])=>entry.endsWith(`:${id}`)&&entry.slice(0,7)<month&&value?.podId).sort(([a],[b])=>a.localeCompare(b)).at(-1)?.[1];return prior?{...person,podId:prior.podId,rowOrder:prior.rowOrder}:person}
export function cellFor(state,id,date){return state.overrides[`${date}:${id}`]||state.assignments[`${date}:${id}`]||(()=>{let member=memberForDate(state,id,date),p=pods.find(x=>x.id===member?.podId);return !p||date<(member?.effectiveFrom||'2026-11-01')?'OFF':cycle(date,p.anchor)})()}
export function initialState(){return {assignments:{},overrides:{}}}
