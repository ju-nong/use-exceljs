import { useRef, useMemo } from "react";
import { Workbook, Worksheet, Borders, FillPattern } from "exceljs";

import { getDay } from "./utils";

type Data = {
    total: number[];
    daily: number[][];
};

const borderStyle: Partial<Borders> = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
};

const fillStyle: FillPattern = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "E8E8E8" },
};

const columnList = [
    "총 제출 문제 수",
    "총 풀이 제출 수",
    "총 사용 선생님 수(유니크)",
    "총 사용 학생 수(유니크)",
];

function stylingTitle(sheet: Worksheet, code: string, label: string) {
    const titleCell = sheet.getCell(code);
    titleCell.value = label;
    titleCell.font = {
        bold: true,
        size: 16,
    };
}

function stylingAllTitle(
    sheet: Worksheet,
    startDate: string,
    endDate: string,
    total: number[],
) {
    sheet.getColumn("A").width = 24;

    stylingTitle(sheet, "A1", "전체");

    sheet.mergeCells("A2:B2");
    const allDateCell = sheet.getCell("A2");
    allDateCell.value = `${startDate} (${getDay(
        startDate,
    )}) ~ ${endDate} (${getDay(endDate)})`;
    allDateCell.border = borderStyle;

    let currentCell = null;
    let tempIndex = 0;
    for (let i = 3; i <= 6; i++) {
        tempIndex = i - 3;

        currentCell = sheet.getCell(`A${i}`);
        currentCell.value = columnList[tempIndex];
        currentCell.border = borderStyle;

        currentCell = sheet.getCell(`B${i}`);
        currentCell.value = total[tempIndex];
        currentCell.fill = fillStyle;
        currentCell.border = borderStyle;
    }
}

function stylingDaily(sheet: Worksheet, startDate: string, daily: number[][]) {
    const dailySize = daily.length;

    if (dailySize > 0) {
        stylingTitle(sheet, "A9", "일별");

        const $startDate = new Date(startDate);
        for (let i = 0; i < dailySize; i++) {
            // sheet.mergeCells("A2:B2");
            // const allDateCell = sheet.getCell("A2");
            // allDateCell.value = `${startDate} (${getDay(
            //     startDate,
            // )}) ~ ${endDate} (${getDay(endDate)})`;
            // allDateCell.border = borderStyle;
        }
    }
}

// Workbook 생성 및 반환
function createWorkbook(
    startDate: string,
    endDate: string,
    { total, daily }: Data,
): Workbook {
    const wb = new Workbook();
    const sheet = wb.addWorksheet(`논술평가사용통계_${startDate}_${endDate}`);

    try {
        // 전체 셀 스타일링
        stylingAllTitle(sheet, startDate, endDate, total);

        // 일별 셀 스타일링
        stylingDaily(sheet, startDate, daily);
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

    // API 응답 예제
    const data = {
        total: [20, 15, 12, 12],
        daily: [
            [3, 3, 1, 1],
            [0, 0, 0, 0],
            [3, 0, 1, 0],
            [1, 2, 1, 2],
            [4, 0, 3, 0],
        ],
    };

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
        const wb = createWorkbook(startDate, endDate, data);

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
