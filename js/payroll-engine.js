class PayrollEngine {
 constructor(rules={}){this.updateRules(rules)}
 updateRules(rules={}){this.rules={mealRate:0,vehicleRate:0,mealLimit:200000,vehicleLimit:200000,npsRate:4.75,nhisRate:3.595,ltciRate:13.14,eiRate:.9,incomeTaxRate:3,localIncomeTaxRate:10,...rules}}
 amount(v){const n=Number(String(v??0).replace(/[,₩원\s]/g,''));return Number.isFinite(n)?Math.max(0,n):0}
 calculateWorkerSalary(worker){
  const r=this.rules,basePay=this.amount(worker.basePay);
  const meal=Math.round(basePay*this.amount(r.mealRate)/100),vehicle=Math.round(basePay*this.amount(r.vehicleRate)/100);
  const overtime=this.amount(worker.overtimePay),loading=this.amount(worker.loadingAllowance),incentive=this.amount(worker.incentive);
  const totalPayment=basePay+meal+vehicle+overtime+loading+incentive;
  const nonTaxableMeal=Math.min(meal,this.amount(r.mealLimit)),nonTaxableVehicle=Math.min(vehicle,this.amount(r.vehicleLimit));
  const taxableIncome=Math.max(0,totalPayment-nonTaxableMeal-nonTaxableVehicle),floor10=v=>Math.floor(v/10)*10;
  const nps=floor10(taxableIncome*this.amount(r.npsRate)/100),nhis=floor10(taxableIncome*this.amount(r.nhisRate)/100);
  const ltci=floor10(nhis*this.amount(r.ltciRate)/100),ei=floor10(taxableIncome*this.amount(r.eiRate)/100);
  const incomeTax=floor10(taxableIncome*this.amount(r.incomeTaxRate)/100),localIncomeTax=floor10(incomeTax*this.amount(r.localIncomeTaxRate)/100);
  const totalDeduction=nps+nhis+ltci+ei+incomeTax+localIncomeTax;
  return {basePay,mealAllowance:meal,vehicleAllowance:vehicle,overtimePay:overtime,loadingAllowance:loading,incentive,nonTaxableMeal,nonTaxableVehicle,taxableIncome,totalPayment,nps,nhis,ltci,ei,incomeTax,localIncomeTax,totalDeduction,netPayment:totalPayment-totalDeduction};
 }
}