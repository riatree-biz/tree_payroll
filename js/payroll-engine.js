class PayrollEngine {
    constructor(rules) {
        this.rules = rules || { mealLimit: 200000, vehicleLimit: 200000 };
    }

    calculateWorkerSalary(worker) {
        // 1. 지급 항목 묶음
        const basePay = Number(worker.basePay) || 0;
        const meal = Number(worker.mealAllowance) || 0;
        const vehicle = Number(worker.vehicleAllowance) || 0;
        const overtime = Number(worker.overtimePay) || 0;
        const loading = Number(worker.loadingAllowance) || 0;
        const incentive = Number(worker.incentive) || 0;

        // 비과세 분리 (식대, 차량운전보조금 등)
        const nonTaxableMeal = Math.min(meal, this.rules.mealLimit);
        const nonTaxableVehicle = Math.min(vehicle, this.rules.vehicleLimit);
        const nonTaxableIncome = nonTaxableMeal + nonTaxableVehicle;

        // 지급합계 (총지급액)
        const totalPayment = basePay + meal + vehicle + overtime + loading + incentive;
        
        // 과세 대상 소득 (지급합계 - 비과세총액)
        const taxableIncome = totalPayment - nonTaxableIncome;

        // 2. 공제 항목 묶음 계산 (4대보험 및 세금)
        const nps = Math.floor((taxableIncome * 0.0475) / 10) * 10;        // 국민연금 (4.75%)
        const nhis = Math.floor((taxableIncome * 0.03595) / 10) * 10;     // 건강보험 (3.595%)
        const ltci = Math.floor((nhis * 0.1314) / 10) * 10;               // 장기요양보험 (건강보험의 13.14%)
        const ei = Math.floor((taxableIncome * 0.009) / 10) * 10;         // 고용보험 (0.9%)
        
        // 소득세 및 지방소득세 간이 계산
        const incomeTax = Math.floor((taxableIncome * 0.03) / 10) * 10; 
        const localIncomeTax = Math.floor((incomeTax * 0.1) / 10) * 10;

        // 공제합계 (총공제액)
        const totalDeduction = nps + nhis + ltci + ei + incomeTax + localIncomeTax;

        // 실지급액
        const netPayment = totalPayment - totalDeduction;

        return {
            basePay, meal, vehicle, overtime, loading, incentive,
            totalPayment,
            nps, nhis, ltci, ei, incomeTax, localIncomeTax,
            totalDeduction,
            netPayment
        };
    }
}