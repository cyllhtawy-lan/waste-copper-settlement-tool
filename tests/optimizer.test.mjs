import assert from 'node:assert/strict';
import {performance} from 'node:perf_hooks';
import {solveAnytime,solveExactSmall} from '../app/optimizer-core.ts';

const rules={auRates:[{min:100,rate:.96},{min:50,rate:.955},{min:30,rate:.95},{min:20,rate:.94},{min:5,rate:.935}],agRates:[{min:2000,rate:.92},{min:1500,rate:.91},{min:1300,rate:.90},{min:1000,rate:.89}]};
const goal={auMin:100,agMin:1000,maxGroup:40};
const raw=[
['广粗63F',26.792,47.34,587.2],['广粗74F',17.847,114.80,1018.2],['广粗78F',23.129,107.29,1082.0],['广粗79F',25.990,110.61,831.8],['广粗80F',25.875,112.70,865.7],['广粗82F',24.560,63.52,1029.9],['广粗100F',26.955,67.95,807.2],['广粗132F',16.619,46.41,687.7],['广粗133F',9.233,122.00,1742.0],['广粗142F',26.839,116.36,1075.3],['广粗143F',26.992,98.24,998.6],['广粗145F',19.466,72.10,863.5],['广粗150F',9.750,56.14,777.4],['广粗151F',12.304,110.07,896.1],['广粗153F',20.944,183.27,2747.4],['广粗158F',26.820,158.51,948.8],['广粗161F',5.064,109.82,860.4],['广粗163F',27.564,58.14,837.4],['广粗215F',1.139,167.49,1465.3],['广粗216F',25.491,140.40,1133.6],['广粗137F',7.345,79.17,3554.7],['广粗138F',19.187,63.52,2881.3],['广粗139F',23.311,76.57,925.3],['广粗152F',20.720,177.43,2842.4],['广粗144F',4.036,95.14,960.8],['广粗67F',24.968,19.00,662.4],['广粗81F',3.726,27.15,900.1],['广粗123F',25.620,45.42,802.2],['广粗128F',26.855,16.11,450.1],['广粗159F',22.761,33.21,783.8],['广粗160F',23.290,34.30,819.0],['广粗162F',21.066,44.42,613.9],['广粗222F',23.996,17.86,556.8],['广粗224F',27.081,12.58,455.6]
].map(([no,tons,au,ag],i)=>({id:String(i),no,tons,au,ag}));

{
 const tiny=raw.slice(0,10),exact=solveExactSmall(tiny,goal,rules),anytime=solveAnytime(tiny,goal,rules,1000);
 assert.equal(anytime.proven,true);assert.equal(anytime.payableAu,exact.au);assert.equal(anytime.payableAg,exact.ag);
}

{
 const started=performance.now(),result=solveAnytime(raw,goal,rules,30000);
 assert.ok(Math.abs(result.payableAu-45717.4269696)<1e-5,`unexpected Au ${result.payableAu}`);
 assert.ok(result.payableAg>=518246.477495-1e-5,`Ag below regression floor: ${result.payableAg}`);
 assert.equal(new Set(result.groups.flat()).size,result.groups.flat().length);
 console.log(JSON.stringify({case:'34-cabinet',milliseconds:Math.round(performance.now()-started),...result},null,2));
}

for(const count of [30,31,40]){
 const sample=Array.from({length:count},(_,i)=>({id:`p${i}`,no:`P${i+1}`,tons:18+i%9,au:70+(i*37)%70,ag:650+(i*211)%900})),started=performance.now(),result=solveAnytime(sample,goal,rules,1000);
 assert.ok(result.groups.flat().length===new Set(result.groups.flat()).size);
 assert.ok(performance.now()-started<1500,`${count} cabinets did not honor the search deadline`);
 console.log(JSON.stringify({case:`performance-${count}`,milliseconds:Math.round(performance.now()-started),covered:result.covered,proven:result.proven}));
}

console.log('optimizer regression tests passed');
