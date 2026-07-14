"""
예실대비표 엑셀들을 읽어 원본 표와 비용 코드 기준 통합 표를 생성합니다.

표 구조(열 위치, 요약 규칙 등)는 LLM이 읽고 금액으로 검증합니다.
LLM 사용 불가 시 규칙 기반 추론으로 자동 전환합니다.

환경변수:
    BUDGET_LLM_MODEL      기본 qwen2.5:7b
    BUDGET_LLM_BASE_URL   기본 http://localhost:11434
    BUDGET_LLM_DISABLED=1 LLM 끄고 규칙 기반만 사용

실행 방법 (프로젝트 루트에서):
    python src/main.py
"""

from pathlib import Path

from reader import collect_all_files, find_excel_files
from writer import save_integrated_workbook

PROJECT_ROOT = Path(__file__).resolve().parent.parent

INPUT_DIR = PROJECT_ROOT / "data" / "input"
OUTPUT_FILE = PROJECT_ROOT / "data" / "output" / "예실대비표_통합결과.xlsx"


def main() -> None:
    """전체 통합 작업을 실행한다."""
    excel_files = find_excel_files(INPUT_DIR, OUTPUT_FILE.name)

    if not excel_files:
        raise FileNotFoundError(
            f"입력 엑셀 파일이 없습니다: {INPUT_DIR.resolve()}"
        )

    print(f"통합 대상 파일 수: {len(excel_files)}개")

    headers, file_tables, integrated_rows, schema = collect_all_files(excel_files)

    if not file_tables or schema is None:
        raise ValueError("정상적으로 처리된 엑셀 파일이 없습니다.")

    save_integrated_workbook(
        headers,
        file_tables,
        integrated_rows,
        OUTPUT_FILE,
    )

    print(f"\n스키마 출처: {schema.source}")
    print(f"탐지된 금액 열 수: {len(headers)}개")
    print(f"원본 표 수: {len(file_tables)}개")
    print(f"결과 시트: {[t['시트명'] for t in file_tables] + ['통합']}")
    print(
        "요약 규칙: "
        + ", ".join(
            f"{rule.label}[{rule.composition}"
            + (f":{'+'.join(rule.codes)}" if rule.codes else "")
            + "]"
            for rule in schema.summary_rules
        )
    )
    print(f"결과 파일: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
