import os
import shutil
from pathlib import Path

import pandas as pd


# ==========================
# 1. 상단 설정
# ==========================

# 엑셀 파일 경로
EXCEL_PATH = Path(r"W:\py\25.11_QeQ2\data_image\지수test\지수test_top10.xlsx")

# A3\A31 하위 전체에서 원본 파일을 찾는다
SEARCH_ROOT = Path(r"W:\py\25.11_QeQ2\data_image\A3\A31")

# 복사 대상 폴더 (없으면 생성)
DEST_DIR = Path(r"W:\py\25.11_QeQ2\web2\candi")

# 읽을 열 인덱스 (0 기반): C, E, G, I, K, M, O, Q, S, U
TARGET_COL_INDEXES = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20]


# ==========================
# 2. 엑셀에서 ID 수집
# ==========================

def collect_ids_from_excel() -> set:
    df = pd.read_excel(EXCEL_PATH, sheet_name=0)

    ids = set()

    for col_idx in TARGET_COL_INDEXES:
        if col_idx >= df.shape[1]:
            continue
        col = df.iloc[:, col_idx]
        for v in col:
            if pd.isna(v):
                continue
            s = str(v).strip()
            if not s:
                continue
            ids.add(s)

    return ids


# ==========================
# 3. 검색용 파일 맵 구축
# ==========================

def build_file_map() -> dict:
    file_map = {}  # key: 파일명, value: 전경로 Path

    for root, dirs, files in os.walk(SEARCH_ROOT):
        root_path = Path(root)
        for name in files:
            if not name.lower().endswith("_p.gif"):
                continue
            full_path = root_path / name
            # 동일 파일명이 여러 개 있으면 처음 것만 사용, 나머지는 경고만
            if name in file_map:
                print(f"[WARN] 중복 파일명 발견 (무시): {full_path}")
                continue
            file_map[name] = full_path

    return file_map


# ==========================
# 4. 메인 로직
# ==========================

def main():
    print("[STEP] 엑셀에서 ID 수집")
    ids = collect_ids_from_excel()
    print(f"[INFO] 엑셀에서 수집한 ID 개수: {len(ids)}")

    if not ids:
        print("[ERROR] 유효한 ID가 없습니다.")
        return

    print("[STEP] 검색 루트 스캔:", SEARCH_ROOT)
    file_map = build_file_map()
    print(f"[INFO] 검색된 *_p.gif 파일 개수: {len(file_map)}")

    DEST_DIR.mkdir(parents=True, exist_ok=True)

    copied = 0
    missing = 0

    for id_val in sorted(ids):
        filename = f"{id_val}_p.gif"
        src = file_map.get(filename)

        if src is None:
            print(f"[MISS] 원본 파일 없음: {filename}")
            missing += 1
            continue

        dest = DEST_DIR / filename
        try:
            shutil.copy2(src, dest)
            copied += 1
        except Exception as e:
            print(f"[ERROR] 복사 실패: {src} -> {dest}  ({e})")

    print("===================================")
    print(f"[DONE] 복사 완료  copied={copied}, missing={missing}")
    print(f"[OUT] 목적지 폴더: {DEST_DIR}")


if __name__ == "__main__":
    main()
