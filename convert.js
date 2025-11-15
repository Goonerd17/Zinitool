/** ===========================================
 * 1) 병합 해제(Unmerge) 기능 (여러 시트 지원)
 * =========================================== */
document.getElementById("fileUnmerge").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const originalName = file.name.replace(/\.xlsx?$/i, ""); // 확장자 제거
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const merges = sheet["!merges"] || [];

        merges.forEach((merge) => {
            const s = merge.s;
            const e2 = merge.e;
            const startCellAddr = XLSX.utils.encode_cell(s);
            const value = sheet[startCellAddr]?.v ?? "";

            for (let r = s.r; r <= e2.r; r++) {
                for (let c = s.c; c <= e2.c; c++) {
                    const addr = XLSX.utils.encode_cell({ r, c });
                    sheet[addr] = { t: "s", v: value };
                }
            }
        });

        sheet["!merges"] = [];
    });

    const out = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${originalName}_unmerged.xlsx`;
    a.click();
});


/** ===========================================
 * 2) 다시 병합 (Re-Merge) 기능 (여러 시트 지원)
 * =========================================== */
document.getElementById("fileMerge").addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const originalName = file.name.replace(/\.xlsx?$/i, "");
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    const colNames = [
        document.getElementById("mergeColumnName1").value.trim(),
        document.getElementById("mergeColumnName2").value.trim(),
        document.getElementById("mergeColumnName3").value.trim(),
    ].filter(n => n);

    if (colNames.length === 0) {
        alert("적어도 하나 이상의 병합할 컬럼명을 입력하세요.");
        return;
    }

    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet["!ref"]);
        const merges = [];

        colNames.forEach(colName => {
            // 1) 헤더 행에서 컬럼 위치 찾기
            let targetCol = -1;
            for (let c = range.s.c; c <= range.e.c; c++) {
                const addr = XLSX.utils.encode_cell({ r: range.s.r, c });
                if ((sheet[addr]?.v ?? "").toString() === colName) {
                    targetCol = c;
                    break;
                }
            }
            if (targetCol === -1) return;

            // 2) 값 기준으로 연속된 구간 병합
            let startRow = range.s.r + 1; // 헤더 다음부터
            let startValue = sheet[XLSX.utils.encode_cell({ r: startRow, c: targetCol })]?.v ?? "";

            for (let r = startRow + 1; r <= range.e.r; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r, c: targetCol });
                const cellValue = sheet[cellAddr]?.v ?? "";

                if (cellValue !== startValue) {
                    if (r - 1 > startRow) {
                        merges.push({
                            s: { r: startRow, c: targetCol },
                            e: { r: r - 1, c: targetCol }
                        });
                    }
                    startRow = r;
                    startValue = cellValue;
                }
            }
            // 마지막 구간 병합
            if (range.e.r > startRow) {
                merges.push({
                    s: { r: startRow, c: targetCol },
                    e: { r: range.e.r, c: targetCol }
                });
            }
        });

        // 3) 병합 적용 & 가운데 정렬
        sheet["!merges"] = merges;
        merges.forEach(m => {
            const s = m.s;
            const e2 = m.e;
            const value = sheet[XLSX.utils.encode_cell(s)]?.v ?? "";

            for (let r = s.r; r <= e2.r; r++) {
                const addr = XLSX.utils.encode_cell({ r, c: s.c });
                sheet[addr] = {
                    t: "s",
                    v: value,
                    s: { alignment: { horizontal: "center", vertical: "center" } }
                };
            }
        });
    });

    const out = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${originalName}_merged.xlsx`;
    a.click();
});


// Unmerge 드래그 & 드롭
const unmergeDropArea = document.getElementById("unmerge-drop-area");
unmergeDropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    unmergeDropArea.style.backgroundColor = "#e0e7ff";
});
unmergeDropArea.addEventListener("dragleave", (e) => {
    e.preventDefault();
    unmergeDropArea.style.backgroundColor = "#f8fafc";
});
unmergeDropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    unmergeDropArea.style.backgroundColor = "#f8fafc";
    if (e.dataTransfer.files.length) {
        document.getElementById("fileUnmerge").files = e.dataTransfer.files;
        document.getElementById("fileUnmerge").dispatchEvent(new Event("change"));
    }
});

// Merge 드래그 & 드롭
const mergeDropArea = document.getElementById("merge-drop-area");
mergeDropArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    mergeDropArea.style.backgroundColor = "#e0e7ff";
});
mergeDropArea.addEventListener("dragleave", (e) => {
    e.preventDefault();
    mergeDropArea.style.backgroundColor = "#f8fafc";
});
mergeDropArea.addEventListener("drop", (e) => {
    e.preventDefault();
    mergeDropArea.style.backgroundColor = "#f8fafc";
    if (e.dataTransfer.files.length) {
        document.getElementById("fileMerge").files = e.dataTransfer.files;
        document.getElementById("fileMerge").dispatchEvent(new Event("change"));
    }
});


