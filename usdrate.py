from __future__ import annotations

import argparse
import csv
import io
import json
import queue
import re
import sys
import threading
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, Y, filedialog, messagebox, ttk
import tkinter as tk
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import urlopen


BASE_URL = "https://oapi.koreaexim.go.kr/site/program/financial/exchangeJSON"
CSV_ENCODINGS = ["utf-8-sig", "utf-8", "cp949", "euc-kr"]
DATE_COLUMN_PRIORITY = {
    "closed": 0,
    "closeddate": 1,
    "date": 2,
    "daydate": 3,
    "day": 4,
    "날짜": 0,
    "거래일": 1,
    "기준일": 2,
    "일자": 3,
}
DATE_VALUE_RE = re.compile(r"\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}|\d{8}")


@dataclass
class ProcessingResult:
    input_encoding: str
    delimiter: str
    date_column: str
    total_rows: int
    parse_failed: int
    rate_missing: int
    output_path: Path


def normalize_column_name(name: str) -> str:
    return re.sub(r"[\W_]+", "", str(name).strip().lower(), flags=re.UNICODE)


def parse_date_text(value: str) -> date | None:
    text = str(value).strip()
    if not text:
        return None

    if re.fullmatch(r"\d{8}", text):
        try:
            return datetime.strptime(text, "%Y%m%d").date()
        except ValueError:
            return None

    normalized = text.replace(".", "-").replace("/", "-")
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", normalized):
        try:
            return datetime.strptime(normalized, "%Y-%m-%d").date()
        except ValueError:
            return None

    return None


def read_csv_with_fallback(path: Path) -> tuple[list[str], list[list[str]], str, csv.Dialect]:
    raw = path.read_bytes()
    failures: list[str] = []

    for encoding in CSV_ENCODINGS:
        try:
            text = raw.decode(encoding)
            dialect = sniff_dialect(text)
            rows = list(csv.reader(io.StringIO(text), dialect))
            if not rows:
                raise ValueError("CSV 파일이 비어 있습니다.")
            headers, data_rows = normalize_rows(rows)
            return headers, data_rows, encoding, dialect
        except Exception as exc:
            failures.append(f"{encoding}: {exc}")

    raise ValueError(
        "CSV를 읽지 못했습니다. 시도한 인코딩: "
        + ", ".join(CSV_ENCODINGS)
        + "\n"
        + "\n".join(failures)
    )


def sniff_dialect(text: str) -> csv.Dialect:
    sample = text[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return csv.excel


def normalize_rows(rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    headers = list(rows[0])
    data_rows = [list(row) for row in rows[1:]]
    max_len = max([len(headers)] + [len(row) for row in data_rows]) if rows else 0

    if max_len == 0:
        raise ValueError("CSV 헤더를 찾지 못했습니다.")

    if len(headers) < max_len:
        for index in range(len(headers), max_len):
            headers.append(f"column_{index + 1}")

    for index, header in enumerate(headers):
        if not str(header).strip():
            headers[index] = f"column_{index + 1}"

    for row in data_rows:
        if len(row) < max_len:
            row.extend([""] * (max_len - len(row)))

    return headers, data_rows


def detect_date_column(headers: list[str], data_rows: list[list[str]]) -> tuple[int, str, list[date | None]]:
    candidates: list[tuple[int, int, int, int, str, list[date | None]]] = []

    for index, header in enumerate(headers):
        parsed_dates = [parse_date_text(row[index]) for row in data_rows]
        parsed_count = sum(value is not None for value in parsed_dates)
        if parsed_count == 0:
            continue

        iso_like_count = sum(bool(DATE_VALUE_RE.search(str(row[index]).strip())) for row in data_rows)
        normalized_name = normalize_column_name(header)

        if normalized_name in DATE_COLUMN_PRIORITY:
            priority = DATE_COLUMN_PRIORITY[normalized_name]
        elif iso_like_count > 0:
            priority = 50
        else:
            priority = 100

        candidates.append((priority, -parsed_count, -iso_like_count, index, header, parsed_dates))

    if not candidates:
        raise ValueError(
            "날짜 컬럼을 찾지 못했습니다. Closed, Date, Day 같은 컬럼이 있거나 "
            "값이 yyyy-mm-dd 형식이어야 합니다."
        )

    candidates.sort(key=lambda item: (item[0], item[1], item[2], item[4]))
    _priority, _parsed_count, _iso_like_count, index, header, parsed_dates = candidates[0]
    return index, header, parsed_dates


def resolve_date_column(
    headers: list[str],
    data_rows: list[list[str]],
    forced_column: str | None,
) -> tuple[int, str, list[date | None]]:
    if not forced_column:
        return detect_date_column(headers, data_rows)

    try:
        index = headers.index(forced_column)
    except ValueError as exc:
        raise ValueError(f"지정한 날짜 컬럼을 찾지 못했습니다: {forced_column}") from exc

    parsed_dates = [parse_date_text(row[index]) for row in data_rows]
    if not any(value is not None for value in parsed_dates):
        raise ValueError(f"지정한 날짜 컬럼에서 날짜 값을 읽지 못했습니다: {forced_column}")
    return index, headers[index], parsed_dates


def fetch_usd_rate_for_ymd(
    ymd: str,
    authkey: str,
    day_cache: dict[str, str | None],
) -> str | None:
    if ymd in day_cache:
        return day_cache[ymd]

    params = urlencode({"authkey": authkey, "searchdate": ymd, "data": "AP01"})
    url = f"{BASE_URL}?{params}"
    last_error: Exception | None = None

    for attempt in range(5):
        try:
            with urlopen(url, timeout=30) as response:
                body = response.read().decode("utf-8")
            data = json.loads(body)
            rate = extract_usd_rate(data)
            day_cache[ymd] = rate
            return rate
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError, ValueError) as exc:
            last_error = exc
            if isinstance(exc, ValueError):
                raise
            if attempt == 4:
                break
            time.sleep(min(2 ** attempt, 5))

    raise RuntimeError(f"{ymd} 환율 조회에 실패했습니다: {last_error}")


def extract_usd_rate(data: object) -> str | None:
    if not isinstance(data, list) or not data:
        return None

    first_row = data[0]
    if isinstance(first_row, dict):
        result_code = str(first_row.get("result", "")).strip()
        if result_code and result_code not in {"1", "01"}:
            message = str(first_row.get("msg", "")).strip() or f"API 오류 코드: {result_code}"
            raise ValueError(message)

    for row in data:
        if not isinstance(row, dict):
            continue
        if str(row.get("cur_unit", "")).strip() != "USD":
            continue
        value = str(row.get("deal_bas_r", "")).replace(",", "").strip()
        return value or None

    return None


def build_fx_map(parsed_dates: list[date | None], authkey: str) -> dict[date, str | None]:
    unique_dates = sorted({value for value in parsed_dates if value is not None})
    fx_map: dict[date, str | None] = {}
    day_cache: dict[str, str | None] = {}

    for original_date in unique_dates:
        lookup_date = original_date
        rate: str | None = None

        for _ in range(14):
            ymd = lookup_date.strftime("%Y%m%d")
            rate = fetch_usd_rate_for_ymd(ymd, authkey, day_cache)
            if rate is not None:
                break
            lookup_date -= timedelta(days=1)

        fx_map[original_date] = rate

    return fx_map


def process_csv(
    input_path: Path,
    output_path: Path,
    authkey: str,
    forced_date_column: str | None = None,
) -> ProcessingResult:
    headers, data_rows, encoding, dialect = read_csv_with_fallback(input_path)
    date_index, date_column, parsed_dates = resolve_date_column(headers, data_rows, forced_date_column)
    fx_map = build_fx_map(parsed_dates, authkey)

    if "환율" in headers:
        rate_index = headers.index("환율")
    else:
        rate_index = date_index + 1
        headers.insert(rate_index, "환율")
        for row in data_rows:
            row.insert(rate_index, "")

    parse_failed = 0
    rate_missing = 0

    for row_index, row in enumerate(data_rows):
        parsed_date = parsed_dates[row_index]
        if parsed_date is None:
            parse_failed += 1
            row[rate_index] = ""
            continue

        rate_value = fx_map.get(parsed_date)
        if rate_value is None:
            rate_missing += 1
            row[rate_index] = ""
        else:
            row[rate_index] = rate_value

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=dialect.delimiter)
        writer.writerow(headers)
        writer.writerows(data_rows)

    return ProcessingResult(
        input_encoding=encoding,
        delimiter=dialect.delimiter,
        date_column=date_column,
        total_rows=len(data_rows),
        parse_failed=parse_failed,
        rate_missing=rate_missing,
        output_path=output_path,
    )


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_환율추가.csv")


def run_cli(args: argparse.Namespace) -> int:
    if not args.input:
        print("입력 파일 경로가 필요합니다. --input 옵션을 입력하세요.", file=sys.stderr)
        return 1

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve() if args.output else default_output_path(input_path)

    if not input_path.is_file():
        print(f"입력 파일을 찾지 못했습니다: {input_path}", file=sys.stderr)
        return 1
    if not args.api_key:
        print("API 키가 필요합니다. --api-key 옵션을 입력하세요.", file=sys.stderr)
        return 1

    try:
        result = process_csv(input_path, output_path, args.api_key, args.date_column)
    except Exception as exc:
        print(f"처리 실패: {exc}", file=sys.stderr)
        return 1

    print("완료되었습니다.")
    print(f"입력 인코딩: {result.input_encoding}")
    print(f"사용한 날짜 컬럼: {result.date_column}")
    print(f"전체 행 수: {result.total_rows}")
    print(f"날짜 해석 실패 행 수: {result.parse_failed}")
    print(f"환율 미조회 행 수: {result.rate_missing}")
    print(f"결과 파일: {result.output_path}")
    return 0


class FxUploaderApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("CSV 환율 추가기")
        self.root.geometry("760x360")
        self.root.minsize(680, 320)

        self.api_key_var = tk.StringVar()
        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.date_column_var = tk.StringVar()
        self.status_var = tk.StringVar(value="CSV 파일과 API 키를 입력한 뒤 실행하세요.")
        self.result_queue: queue.Queue[tuple[str, object]]
        self.result_queue = queue.Queue()
        self.is_running = False

        self.build_ui()

    def build_ui(self) -> None:
        root_frame = ttk.Frame(self.root, padding=16)
        root_frame.pack(fill=BOTH, expand=True)

        ttk.Label(root_frame, text="입력 CSV").pack(anchor="w")
        input_frame = ttk.Frame(root_frame)
        input_frame.pack(fill=X, pady=(4, 12))
        ttk.Entry(input_frame, textvariable=self.input_path_var).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(input_frame, text="찾기", command=self.choose_input).pack(side=RIGHT, padx=(8, 0))

        ttk.Label(root_frame, text="저장 CSV").pack(anchor="w")
        output_frame = ttk.Frame(root_frame)
        output_frame.pack(fill=X, pady=(4, 12))
        ttk.Entry(output_frame, textvariable=self.output_path_var).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(output_frame, text="저장 위치", command=self.choose_output).pack(side=RIGHT, padx=(8, 0))

        ttk.Label(root_frame, text="API 키").pack(anchor="w")
        ttk.Entry(root_frame, textvariable=self.api_key_var, show="*").pack(fill=X, pady=(4, 12))

        ttk.Label(root_frame, text="날짜 컬럼명 (비우면 자동 탐지)").pack(anchor="w")
        ttk.Entry(root_frame, textvariable=self.date_column_var).pack(fill=X, pady=(4, 12))

        button_frame = ttk.Frame(root_frame)
        button_frame.pack(fill=X, pady=(4, 12))
        self.run_button = ttk.Button(button_frame, text="실행", command=self.start_processing)
        self.run_button.pack(side=LEFT)
        ttk.Button(button_frame, text="닫기", command=self.root.destroy).pack(side=LEFT, padx=(8, 0))

        status_frame = ttk.LabelFrame(root_frame, text="상태", padding=12)
        status_frame.pack(fill=BOTH, expand=True)

        self.status_text = tk.Text(status_frame, height=8, wrap="word")
        self.status_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        self.log(self.status_var.get())

    def log(self, message: str) -> None:
        self.status_var.set(message)
        self.status_text.insert(END, f"{message}\n")
        self.status_text.see(END)

    def choose_input(self) -> None:
        filename = filedialog.askopenfilename(
            title="입력 CSV 선택",
            filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
        )
        if not filename:
            return

        input_path = Path(filename)
        self.input_path_var.set(str(input_path))

        current_output = self.output_path_var.get().strip()
        if not current_output:
            self.output_path_var.set(str(default_output_path(input_path)))

    def choose_output(self) -> None:
        initial_name = "result.csv"
        input_value = self.input_path_var.get().strip()
        if input_value:
            initial_name = default_output_path(Path(input_value)).name

        filename = filedialog.asksaveasfilename(
            title="결과 CSV 저장 위치",
            defaultextension=".csv",
            initialfile=initial_name,
            filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
        )
        if filename:
            self.output_path_var.set(filename)

    def start_processing(self) -> None:
        if self.is_running:
            return

        input_value = self.input_path_var.get().strip()
        output_value = self.output_path_var.get().strip()
        authkey = self.api_key_var.get().strip()
        forced_column = self.date_column_var.get().strip() or None

        if not input_value:
            messagebox.showerror("입력 오류", "입력 CSV 파일을 선택하세요.")
            return
        if not authkey:
            messagebox.showerror("입력 오류", "API 키를 입력하세요.")
            return

        input_path = Path(input_value).expanduser().resolve()
        if not input_path.is_file():
            messagebox.showerror("입력 오류", f"입력 파일을 찾지 못했습니다.\n{input_path}")
            return

        output_path = Path(output_value).expanduser().resolve() if output_value else default_output_path(input_path)
        self.output_path_var.set(str(output_path))

        self.is_running = True
        self.run_button.state(["disabled"])
        self.log("처리 중입니다. 파일 크기와 날짜 수에 따라 시간이 걸릴 수 있습니다.")

        thread = threading.Thread(
            target=self.process_in_background,
            args=(input_path, output_path, authkey, forced_column),
            daemon=True,
        )
        thread.start()
        self.root.after(150, self.poll_result_queue)

    def process_in_background(
        self,
        input_path: Path,
        output_path: Path,
        authkey: str,
        forced_column: str | None,
    ) -> None:
        try:
            result = process_csv(input_path, output_path, authkey, forced_column)
        except Exception as exc:
            self.result_queue.put(("error", exc))
            return

        self.result_queue.put(("success", result))

    def poll_result_queue(self) -> None:
        try:
            state, payload = self.result_queue.get_nowait()
        except queue.Empty:
            if self.is_running:
                self.root.after(150, self.poll_result_queue)
            return

        self.is_running = False
        self.run_button.state(["!disabled"])

        if state == "error":
            message = f"처리 실패: {payload}"
            self.log(message)
            messagebox.showerror("실패", message)
            return

        result = payload
        assert isinstance(result, ProcessingResult)

        self.log("완료되었습니다.")
        self.log(f"입력 인코딩: {result.input_encoding}")
        self.log(f"사용한 날짜 컬럼: {result.date_column}")
        self.log(f"날짜 해석 실패 행 수: {result.parse_failed}")
        self.log(f"환율 미조회 행 수: {result.rate_missing}")
        self.log(f"결과 파일: {result.output_path}")

        summary = "\n".join(
            [
                "완료되었습니다.",
                f"입력 인코딩: {result.input_encoding}",
                f"사용한 날짜 컬럼: {result.date_column}",
                f"날짜 해석 실패 행 수: {result.parse_failed}",
                f"환율 미조회 행 수: {result.rate_missing}",
                f"결과 파일: {result.output_path}",
            ]
        )
        messagebox.showinfo("완료", summary)

    def run(self) -> None:
        self.root.mainloop()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="CSV 날짜 기준으로 USD 환율 컬럼을 추가합니다.")
    parser.add_argument("-i", "--input", help="입력 CSV 파일 경로")
    parser.add_argument("-o", "--output", help="출력 CSV 파일 경로")
    parser.add_argument("--api-key", help="한국수출입은행 API 키")
    parser.add_argument("--date-column", help="날짜 컬럼명 직접 지정")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.input or args.output or args.api_key or args.date_column:
        return run_cli(args)

    app = FxUploaderApp()
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
