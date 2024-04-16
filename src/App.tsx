function createURL(): string {
    return window.URL.createObjectURL(
        new Blob(
            ["Hello World"],

            {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
        ),
    );
}

function createLink(url: string): HTMLAnchorElement {
    const link = document.createElement("a");
    link.href = url;
    link.download = "Exam";

    return link;
}

function App() {
    function handleDownload() {
        const url = createURL();
        const link = createLink(url);

        document.body.appendChild(link);
        link.click();
        link.remove();

        window.URL.revokeObjectURL(url);
    }

    return (
        <div>
            <button onClick={handleDownload}>엑셀 다운로드</button>
        </div>
    );
}

export default App;
