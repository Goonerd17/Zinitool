/** ===========================================
 * 1) 병합 해제(Unmerge) 기능
 * =========================================== */
document.getElementById("fileUnmerge").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const merges = sheet["!merges"] || [];

    // 1) 병합된 셀 해제 & 값 복사
    merges.forEach((merge) => {
        const s = merge.s; // start
        const e2 = merge.e; // end

        const startCellAddr = XLSX.utils.encode_cell(s);
        const value = sheet[startCellAddr]?.v ?? "";

        for (let r = s.r; r <= e2.r; r++) {
            for (let c = s.c; c <= e2.c; c++) {
                const addr = XLSX.utils.encode_cell({ r, c });
                sheet[addr] = { t: "s", v: value };
            }
        }
    });

    // 병합 정보 제거
    sheet["!merges"] = [];

    // 2) 다운로드
    const out = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "converted_unmerged.xlsx";
    a.click();
});


/** ===========================================
 * 2) 다시 병합 (Re-Merge)
 * =========================================== */
document.getElementById("fileMerge").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const merges = [];

    /** 
     * 기본 규칙:
     *  - 같은 열(col)을 기준으로
     *  - 연속된 행 값이 같으면 병합
     *  - 기본적으로 2번째 열(col = 1)을 기준으로 병합
     */
    const targetCol = 1; // 두 번째 열

    let startRow = range.s.r;
    let currentValue = sheet[XLSX.utils.encode_cell({ r: startRow, c: targetCol })]?.v;

    for (let r = range.s.r + 1; r <= range.e.r; r++) {
        const cellAddr = XLSX.utils.encode_cell({ r, c: targetCol });
        const cellValue = sheet[cellAddr]?.v;

        if (cellValue !== currentValue) {
            if (r - 1 > startRow) {
                merges.push({
                    s: { r: startRow, c: targetCol },
                    e: { r: r - 1, c: targetCol }
                });
            }
            startRow = r;
            currentValue = cellValue;
        }
    }

    // 마지막 구간 병합
    if (range.e.r > startRow) {
        merges.push({
            s: { r: startRow, c: targetCol },
            e: { r: range.e.r, c: targetCol }
        });
    }

    // 병합 정보 적용
    sheet["!merges"] = merges;

    // 2) 다운로드
    const out = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "converted_merged.xlsx";
    a.click();
});
