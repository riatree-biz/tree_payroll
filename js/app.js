
document.addEventListener('DOMContentLoaded', () => {
    let currentWorkbook = null;
    let workersList = [];
    let isCalculated = false;

    const payrollEngine = new PayrollEngine();

    const fileInput = document.getElementById('excel-file-input');
    const sheetSelect = document.getElementById('excel-sheet-select');
    const sheetSelector = document.getElementById('excel-sheet-selector');
    const preview = document.getElementById('excel-preview-container');
    const tbody = document.getElementById('payroll-tbody');

    // =====================================================
    // 공통 유틸리티
    // =====================================================

    function normalizeHeader(value) {
        return String(value ?? '')
            .trim()
            .toLowerCase()
            .replace(/[\s_-]+/g, '');
    }

    function toNumber(value) {
        if (value === null || value === undefined || value === '') return 0;
        if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
        if (value instanceof Date) return 0;

        let cleaned = String(value)
            .replace(/,/g, '')
            .replace(/[₩원\s]/g, '')
            .trim();

        if (!cleaned) return 0;

        const negative = /^\((.*)\)$/.exec(cleaned);
        if (negative) cleaned = `-${negative[1]}`;

        const number = Number(cleaned);
        return Number.isFinite(number) ? number : 0;
    }

    function formatMoney(value) {
        return toNumber(value).toLocaleString('ko-KR');
    }

    function showError(message, error) {
        console.error(message, error || '');

        alert(
            message +
            (error?.message ? '\n상세 내용: ' + error.message : '')
        );
    }

    // =====================================================
    // 탭 전환
    // =====================================================

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn')
                .forEach(item => item.classList.remove('active'));

            document.querySelectorAll('.tab-content')
                .forEach(item => item.classList.remove('active'));

            btn.classList.add('active');

            const targetId = btn.getAttribute('data-target');
            const target = document.getElementById(targetId);

            if (target) {
                target.classList.add('active');
            }
        });
    });

    // =====================================================
    // 엑셀 파일 업로드
    // =====================================================

    if (!fileInput || !sheetSelect || !sheetSelector) {
        console.error('엑셀 업로드 관련 HTML 요소를 찾을 수 없습니다.');
        return;
    }

    fileInput.addEventListener('change', async event => {
        const file = event.target.files?.[0];

        if (!file) return;

        try {
            // 새 파일을 읽기 전에 기존 데이터 초기화
            currentWorkbook = null;
            workersList = [];
            isCalculated = false;

            sheetSelect.replaceChildren();

            if (preview) {
                preview.replaceChildren();
                preview.style.display = 'none';
            }

            sheetSelector.style.display = 'none';

            const result = await ExcelHandler.parseExcelFile(file);

            if (
                !result ||
                !result.workbook ||
                !Array.isArray(result.sheetNames)
            ) {
                throw new Error(
                    '엑셀 파일을 읽지 못했습니다. 파일 형식을 확인해주세요.'
                );
            }

            currentWorkbook = result.workbook;

            // 모든 시트 목록 표시
            result.sheetNames.forEach(sheetName => {
                const option = document.createElement('option');

                option.value = sheetName;
                option.textContent = sheetName;

                sheetSelect.appendChild(option);
            });

            sheetSelector.style.display =
                result.sheetNames.length > 0 ? 'block' : 'none';

            // 시트별 실제 사용 범위 확인
            console.group('엑셀 파일 로드 결과');
            console.log('파일명:', file.name);
            console.log('파일 크기:', file.size);
            console.log('시트 목록:', result.sheetNames);

            result.sheetNames.forEach(sheetName => {
                const sheet = currentWorkbook.Sheets[sheetName];

                console.log(
                    sheetName,
                    '사용 범위:',
                    sheet?.['!ref'] || '데이터 없음'
                );
            });

            console.groupEnd();

            if (result.sheetNames.length === 0) {
                throw new Error('엑셀 파일에 시트가 없습니다.');
            }

            if (preview) {
                preview.style.display = 'block';

                const message = document.createElement('p');
                message.textContent =
                    `${file.name} 파일을 읽었습니다. ` +
                    `시트 ${result.sheetNames.length}개를 찾았습니다.`;

                preview.replaceChildren(message);
            }
        } catch (error) {
            currentWorkbook = null;
            workersList = [];

            sheetSelector.style.display = 'none';

            showError('엑셀 파일 로드에 실패했습니다.', error);
        }
    });

    // =====================================================
    // 엑셀 데이터 추출
    //
    // 중요:
    // 실제 데이터 추출은 ExcelHandler.getSheetData()에서
    // 수행합니다. 아래 코드는 반환된 직원 데이터를 검증합니다.
    // =====================================================

    const parseButton = document.getElementById('btn-parse-sheet');

    if (parseButton) {
        parseButton.addEventListener('click', () => {
            if (!currentWorkbook) {
                alert('먼저 엑셀 파일을 업로드해주세요.');
                return;
            }

            const sheetName = sheetSelect.value;
            const worksheet = currentWorkbook.Sheets[sheetName];

            if (!worksheet) {
                alert('올바른 시트를 선택해주세요.');
                return;
            }

            try {
                // 선택한 시트의 데이터를 가져옵니다.
                const extractedData = ExcelHandler.getSheetData(
                    currentWorkbook,
                    sheetName
                );

                if (!Array.isArray(extractedData)) {
                    throw new Error(
                        'getSheetData()가 배열을 반환하지 않았습니다.'
                    );
                }

                // 기존 급여 계산 결과가 새 데이터에 섞이지 않도록 복사
                workersList = extractedData.map(worker => ({ ...worker }));
                isCalculated = false;

                // 추출 결과 및 중복 여부 확인
                console.group(`시트 분석: ${sheetName}`);

                console.log('시트 사용 범위:', worksheet['!ref']);
                console.log('추출된 직원 수:', workersList.length);
                console.log('첫 번째 직원:', workersList[0] ?? null);
                console.log('원본 행 번호:', workersList.map(worker => worker.sourceRow));
                console.log('두 번째 직원:', workersList[1] ?? null);
                console.log('전체 추출 결과:');

                console.table(workersList);

                // 직원 이름이 같은 행 확인
                const nameCounts = new Map();

                workersList.forEach(worker => {
                    const name = String(worker.name ?? '').trim();

                    if (name) {
                        nameCounts.set(
                            name,
                            (nameCounts.get(name) || 0) + 1
                        );
                    }
                });

                const duplicatedNames = [...nameCounts.entries()]
                    .filter(([, count]) => count > 1);

                if (duplicatedNames.length > 0) {
                    console.warn(
                        '동일한 이름이 여러 번 나타납니다. 동명이인일 수 있습니다.',
                        duplicatedNames
                    );
                }

                console.groupEnd();

                // 미리보기 표시
                renderImportSummary(sheetName, workersList);

                // 급여대장 탭으로 이동
                document.querySelector(
                    '[data-target="tab-payroll"]'
                )?.click();

                renderPayrollTable(workersList, false);

            } catch (error) {
                showError('엑셀 시트 데이터 추출에 실패했습니다.', error);
            }
        });
    }

    // =====================================================
    // 데이터 미리보기 요약
    // =====================================================

    function renderImportSummary(sheetName, data) {
        if (!preview) return;

        preview.replaceChildren();
        preview.style.display = 'block';

        const heading = document.createElement('h3');
        heading.textContent =
            `업로드된 데이터 미리보기 (${data.length}명)`;

        const sheetInfo = document.createElement('p');
        sheetInfo.textContent = `선택한 시트: ${sheetName}`;

        const warningInfo = document.createElement('p');

        const missingNameRows = data
            .map((worker, index) => ({
                row: worker?.sourceRow ?? index + 2,
                name: worker?.name
            }))
            .filter(item =>
                !String(item.name ?? '').trim()
            );

        warningInfo.textContent = missingNameRows.length
            ? `이름이 비어 있는 데이터 ${missingNameRows.length}건`
            : '이름 누락 검사 완료';

        const instruction = document.createElement('p');
        instruction.textContent =
            '엑셀의 실제 헤더명을 기준으로 직원 행을 읽었습니다. 원본 행 번호도 함께 보존됩니다.';

        const dataCheck = document.createElement('p');
        const invalidPayRows = data.filter(worker =>
            !Number.isFinite(Number(worker.basePay)) || Number(worker.basePay) < 0
        );

        console.log('기본급 값 확인이 필요한 데이터:', invalidPayRows);
        dataCheck.textContent = invalidPayRows.length
            ? `기본급 값 확인이 필요한 데이터 ${invalidPayRows.length}건`
            : '기본급 데이터 검사 완료';

        preview.append(
            heading,
            sheetInfo,
            warningInfo,
            dataCheck,
            instruction
        );
    }

    // =====================================================
    // 급여 계산
    // =====================================================

    function runCalculation() {
        if (!workersList.length) {
            alert('데이터가 없습니다. 엑셀을 먼저 업로드해주세요.');
            return;
        }

        const invalidRows = workersList
            .map((worker, index) => ({
                row: worker?.sourceRow ?? index + 1,
                worker
            }))
            .filter(({ worker }) =>
                !worker ||
                typeof worker !== 'object' ||
                !String(worker.name ?? '').trim()
            );

        if (invalidRows.length > 0) {
            console.warn('이름이 없거나 잘못된 데이터:', invalidRows);

            const confirmed = confirm(
                `이름이 비어 있거나 잘못된 데이터가 ` +
                `${invalidRows.length}건 있습니다.\n` +
                '계속 계산하시겠습니까?'
            );

            if (!confirmed) return;
        }

        try {
            // 재계산할 때 이전 계산 결과가 중복 적용되지 않도록
            // 원본 추출 데이터에서 다시 계산합니다.
            const sourceData = workersList.map(worker => ({ ...worker }));

            const calculatedData = sourceData.map(worker => {
                const salaryResult =
                    payrollEngine.calculateWorkerSalary(worker);

                if (
                    !salaryResult ||
                    typeof salaryResult !== 'object'
                ) {
                    throw new Error(
                        `${worker.name || '이름 없는 직원'}의 ` +
                        '급여 계산 결과가 올바르지 않습니다.'
                    );
                }

                return {
                    ...worker,
                    ...salaryResult
                };
            });

            workersList = calculatedData;
            isCalculated = true;

            console.group('급여 계산 결과');
            console.log('계산된 직원 수:', workersList.length);
            console.table(workersList);
            console.groupEnd();

            renderPayrollTable(workersList, true);

            alert('지급 및 공제 내역 계산이 완료되었습니다.');

        } catch (error) {
            showError('급여 계산 중 오류가 발생했습니다.', error);
        }
    }

    document.getElementById('footer-btn-calc')
        ?.addEventListener('click', runCalculation);

    document.getElementById('btn-calculate-all')
        ?.addEventListener('click', runCalculation);

    // =====================================================
    // 안전한 테이블 셀 생성
    // =====================================================

    function addCell(row, value, className = '') {
        const cell = document.createElement('td');

        if (className) {
            cell.className = className;
        }
        
        cell.textContent = value == null ? '' : String(value);

        row.appendChild(cell);

        return cell;
    }

    function addMoneyCell(row, value, options = {}) {
        const cell = addCell(row, formatMoney(value), 'number');

        if (options.color) {
            cell.style.color = options.color;
        }

        if (options.strong) {
            const strong = document.createElement('strong');
            strong.textContent = formatMoney(value);
            cell.replaceChildren(strong);
        }

        return cell;
    }

    // =====================================================
    // 급여대장 테이블 렌더링
    // =====================================================

    function renderPayrollTable(data, calculated) {
        if (!tbody) {
            console.error('payroll-tbody 요소를 찾을 수 없습니다.');
            return;
        }

        tbody.replaceChildren();

        data.forEach(worker => {
            const row = document.createElement('tr');

            const nameCell = addCell(row, '');
            const nameStrong = document.createElement('strong');
            nameStrong.textContent = worker.name ?? '';
            nameCell.appendChild(nameStrong);

            addCell(row, worker.dept ?? '');
            
            worker.basePay = toNumber(Math.round(worker.basePay));
            addMoneyCell(row, worker.basePay);
            addMoneyCell(row, worker.mealAllowance);
            addMoneyCell(row, worker.vehicleAllowance);
            addMoneyCell(row, worker.overtimePay);
            addMoneyCell(row, worker.loadingAllowance);
            addMoneyCell(row, worker.incentive);

            const totalPayment = calculated
                ? worker.totalPayment
                : toNumber(worker.basePay) +
                  toNumber(worker.mealAllowance) +
                  toNumber(worker.vehicleAllowance) +
                  toNumber(worker.overtimePay) +
                  toNumber(worker.loadingAllowance) +
                  toNumber(worker.incentive);

            addMoneyCell(row, totalPayment, { strong: true });

            addMoneyCell(row, calculated ? worker.nps : 0);
            addMoneyCell(row, calculated ? worker.nhis : 0);
            addMoneyCell(row, calculated ? worker.ltci : 0);
            addMoneyCell(row, calculated ? worker.ei : 0);
            addMoneyCell(row, calculated ? worker.incomeTax : 0);
            addMoneyCell(row, calculated ? worker.localIncomeTax : 0);

            addMoneyCell(
                row,
                calculated ? worker.totalDeduction : 0,
                { color: '#dc2626', strong: true }
            );

            addMoneyCell(
                row,
                calculated ? worker.netPayment : 0,
                { color: '#1a56db', strong: true }
            );

            tbody.appendChild(row);
        });
    }
});
