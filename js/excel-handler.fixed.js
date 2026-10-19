class ExcelHandler {
    static async parseExcelFile(file) {
        if (!file) throw new Error('엑셀 파일이 선택되지 않았습니다.');
        const extension = file.name.split('.').pop()?.toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(extension)) {
            throw new Error('지원하지 않는 파일 형식입니다. xlsx, xls, csv 파일만 사용할 수 있습니다.');
        }

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error('엑셀 파일을 읽는 중 오류가 발생했습니다.'));
            reader.onload = event => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, {
                        type: 'array',
                        cellDates: true,
                        cellNF: false,
                        cellText: false
                    });
                    if (!workbook.SheetNames?.length) throw new Error('엑셀 파일에 시트가 없습니다.');
                    resolve({ workbook, sheetNames: workbook.SheetNames });
                } catch (error) { reject(error); }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    static normalizeHeader(value) {
        return String(value ?? '')
            .replace(/\uFEFF/g, '')
            .replace(/\r?\n/g, '')
            .replace(/\s+/g, '')
            .replace(/[_-]/g, '')
            .trim()
            .toLowerCase();
    }

    static normalizeText(value) {
        if (value === null || value === undefined) return '';
        return String(value).replace(/\uFEFF/g, '').replace(/\r?\n/g, ' ').trim();
    }

    static toNumber(value, defaultValue = 0) {
        if (value === null || value === undefined || value === '') return defaultValue;
        if (typeof value === 'number') return Number.isFinite(value) ? value : defaultValue;
        if (value instanceof Date) return defaultValue;

        let text = String(value).replace(/,/g, '').replace(/[₩원\s]/g, '').trim();
        if (!text) return defaultValue;
        const negative = /^\((.*)\)$/.exec(text);
        if (negative) text = `-${negative[1]}`;
        const number = Number(text);
        return Number.isFinite(number) ? number : defaultValue;
    }

    static findHeaderRow(rows) {
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const normalized = rows[rowIndex].map(cell => this.normalizeHeader(cell));
            if (normalized.includes('사원명') && normalized.includes('부서명') && normalized.includes('기본급')) {
                return rowIndex;
            }
        }
        return -1;
    }

    static createHeaderMap(headerRow) {
        const map = new Map();
        headerRow.forEach((header, index) => {
            const key = this.normalizeHeader(header);
            if (key && !map.has(key)) map.set(key, index);
        });
        return map;
    }

    static findColumn(headerMap, aliases) {
        for (const alias of aliases) {
            const key = this.normalizeHeader(alias);
            if (headerMap.has(key)) return headerMap.get(key);
        }
        return -1;
    }

    static getCell(row, headerMap, aliases) {
        const index = this.findColumn(headerMap, aliases);
        return index >= 0 ? row[index] : null;
    }

    static getSheetData(workbook, sheetName) {
        const sheet = workbook?.Sheets?.[sheetName];
        if (!sheet) throw new Error(`시트 "${sheetName}"을 찾을 수 없습니다.`);

        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: true,
            defval: null,
            blankrows: true
        });
        if (!rows.length) return [];

        const headerRowIndex = this.findHeaderRow(rows);
        if (headerRowIndex < 0) {
            throw new Error('급여대장 헤더를 찾지 못했습니다. "사원명", "부서명", "기본급" 열을 확인해주세요.');
        }

        const headerMap = this.createHeaderMap(rows[headerRowIndex]);
        const columns = {
            no: ['No', '번호'], name: ['사원명', '성명', '직원명', '이름'],
            dept: ['부서명', '부서'], position: ['직위', '직급'], employeeNo: ['사원번호', '직원번호', '사번'],
            basePay: ['기본급'], mealAllowance: ['식대', '식비', '식사수당'],
            vehicleAllowance: ['차량운전보조금', '차량운전보조', '차량보조금'],
            overtimePay: ['연장근로수당', '연장수당'], nightPay: ['야간근로수당', '야간수당'],
            holidayPay: ['휴일근로수당', '휴일수당'], loadingAllowance: ['상하차 수당', '상하차수당'],
            annualLeavePay: ['연차수당'], incentive: ['인센티브'], totalPayment: ['지급합계'], taxableTotal: ['과세합계'],
            nps: ['국민연금'], nhis: ['건강보험'], ltci: ['장기요양보험'], ei: ['고용보험'],
            incomeTax: ['소득세'], localIncomeTax: ['지방소득세'], yearEndIncomeTax: ['연말소득세'],
            yearEndLocalIncomeTax: ['연말지방소득세'], totalDeduction: ['공제합계'], netPayment: ['실 지급액', '실지급액'],
            workDays: ['근로일수'], totalWorkHours: ['총 근로시간'], overtimeHours: ['연장근로시간'], ordinaryHourlyRate: ['통상시급']
        };

        for (const key of ['name', 'dept', 'basePay']) {
            if (this.findColumn(headerMap, columns[key]) < 0) throw new Error(`필수 열을 찾지 못했습니다: ${key}`);
        }

        const employees = [];
        for (let rowIndex = headerRowIndex + 1; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            if (!row.some(value => value !== null && value !== undefined && String(value).trim() !== '')) continue;

            const name = this.normalizeText(this.getCell(row, headerMap, columns.name));
            // 실제 직원 행만 가져옵니다. 이름이 없으면 합계/설명 영역으로 간주합니다.
            if (!name) continue;

            const num = key => this.toNumber(this.getCell(row, headerMap, columns[key]));
            const text = key => this.normalizeText(this.getCell(row, headerMap, columns[key]));

            employees.push({
                sourceRow: rowIndex + 1,
                no: this.getCell(row, headerMap, columns.no),
                name,
                dept: text('dept'),
                position: text('position'),
                employeeNo: text('employeeNo'),
                workDays: num('workDays'), totalWorkHours: num('totalWorkHours'),
                overtimeHours: num('overtimeHours'), ordinaryHourlyRate: num('ordinaryHourlyRate'),
                basePay: num('basePay'), mealAllowance: num('mealAllowance'), vehicleAllowance: num('vehicleAllowance'),
                overtimePay: num('overtimePay'), nightPay: num('nightPay'), holidayPay: num('holidayPay'),
                loadingAllowance: num('loadingAllowance'), annualLeavePay: num('annualLeavePay'), incentive: num('incentive'),
                excelTotalPayment: num('totalPayment'), excelTaxableTotal: num('taxableTotal'),
                excelNps: num('nps'), excelNhis: num('nhis'), excelLtci: num('ltci'), excelEi: num('ei'),
                excelIncomeTax: num('incomeTax'), excelLocalIncomeTax: num('localIncomeTax'),
                excelTotalDeduction: num('totalDeduction'), excelNetPayment: num('netPayment')
            });
        }

        if (!employees.length) throw new Error(`헤더는 찾았지만 직원 데이터를 찾지 못했습니다. (헤더 행: ${headerRowIndex + 1}행)`);
        return employees;
    }
}
