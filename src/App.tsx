import { useRef, useMemo } from "react";
import { Workbook } from "exceljs";

// Workbook 생성 및 반환
function createWorkbook(startDate: string, endDate: string): Workbook {
    const wb = new Workbook();
    const sheet = wb.addWorksheet(`논술평가사용통계_${startDate}_${endDate}`);

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
