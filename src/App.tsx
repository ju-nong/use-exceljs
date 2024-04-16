import { useRef, useMemo } from "react";
import { Workbook, Worksheet, Borders } from "exceljs";

const borderStyle: Partial<Borders> = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
};

const columnList = [
    "총 제출 문제 수",
    "총 풀이 제출 수",
    "총 사용 선생님 수(유니크)",
    "총 사용 학생 수(유니크)",
];

function stylingAllTitle(sheet: Worksheet, startDate: string, endDate: string) {
    sheet.getColumn("A").width = 22;

    const allTitleCell = sheet.getCell("A1");

    allTitleCell.value = "전체";
    allTitleCell.font = {
        bold: true,
        size: 16,
    };

    sheet.mergeCells("A2:B2");
    const allDateCell = sheet.getCell("A2");
    allDateCell.value = `${startDate} ~ ${endDate}`;
    allDateCell.border = {
        ...borderStyle,
    };

    let currentCell = null;
    for (let i = 3; i <= 6; i++) {
        currentCell = sheet.getCell(`A${i}`);
        currentCell.value = columnList[i - 3];
    }
}

// Workbook 생성 및 반환
function createWorkbook(startDate: string, endDate: string): Workbook {
    const wb = new Workbook();
    const sheet = wb.addWorksheet(`논술평가사용통계_${startDate}_${endDate}`);

    try {
        // 전체 셀 스타일링
        stylingAllTitle(sheet, startDate, endDate);
    } catch (error) {
        console.log(error);
    }

    return wb;
}

// 엑셀 파일 정보 생성 및 반환
async function createURL(workbook: Workbook): Promise<string> {
    const data = await workbook.xlsx.writeBuffer();

    return window.URL.createObjectURL(
        new Blob([data], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
    );
}

function createLink(
    url: string,
    startDate: string,
    endDate: string,
): HTMLAnchorElement {
    const link = document.createElement("a");
    link.href = url;
    link.download = `논술평가사용통계_${startDate}_${endDate}`;

    return link;
}

function App() {
    const $startDate = useRef<HTMLInputElement>(null);
    const $endDate = useRef<HTMLInputElement>(null);

    async function handleDownload() {
        if ($startDate.current === null || $endDate.current === null) {
            return;
        }

        const { value: startDate } = $startDate.current;
        const { value: endDate } = $endDate.current;

        if (startDate.length < 1 || endDate.length < 1) {
            alert("날짜를 입력해주세요.");
            return;
        }

        // Workbook 생성
        const wb = createWorkbook(startDate, endDate);

        // Excel Object URL 생성
        const url = await createURL(wb);

        // Download Link 생성
        const link = createLink(url, startDate, endDate);

        document.body.appendChild(link);
        link.click();
        link.remove();

        window.URL.revokeObjectURL(url);
    }

    return (
        <div>
            Start date
            <input type="date" ref={$startDate} /> <br />
            <br />
            End date
            <input type="date" ref={$endDate} /> <br />
            <br />
            <button onClick={handleDownload}>엑셀 다운로드</button>
        </div>
    );
}

export default App;
