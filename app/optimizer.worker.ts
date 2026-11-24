import {solveAnytime,type OptimizerCabinet,type OptimizerGoal,type OptimizerRules} from './optimizer-core';

self.onmessage=(event:MessageEvent<{jobId:number;items:OptimizerCabinet[];goal:OptimizerGoal;rules:OptimizerRules;timeLimitMs:number}>)=>{
 const{jobId,items,goal,rules,timeLimitMs}=event.data;
 try{const solution=solveAnytime(items,goal,rules,timeLimitMs,x=>self.postMessage({jobId,type:'progress',solution:x}));self.postMessage({jobId,type:'done',solution})}
 catch(error){self.postMessage({jobId,type:'error',error:error instanceof Error?error.message:String(error)})}
};
