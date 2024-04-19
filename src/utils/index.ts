const days = ["일", "월", "화", "수", "목", "금", "토"];

function getDay(date: string) {
    const $date = new Date(date);

    return days[$date.getDay()];
}

export { getDay };
