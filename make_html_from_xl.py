import os
from pathlib import Path

import pandas as pd

# ==========================
# 1. 상단 설정
# ==========================

# Top10 엑셀 경로 (지수test_top10.xlsx)
EXCEL_PATH = Path(r"W:\py\25.11_QeQ2\data_image\지수test\지수test_top10.xlsx")

# 시트 (이름 또는 0,1,2 인덱스)
SHEET_NAME = 0

# A열: 샘플(원본) 파일명
SAMPLE_COL = "sample_filename"   # 헤더 이름이 다르면 바꿔줘

# 1~10등 점수/ID 컬럼 이름
#   B: score1, C: id1, D: score2, E: id2, ...
SCORE_COLS = [f"score{i}" for i in range(1, 11)]
ID_COLS    = [f"id{i}"    for i in range(1, 11)]

# --------------- 이미지 폴더 / 파일명 패턴 ---------------

# A열(샘플) 이미지들이 모여 있는 폴더
SAMPLE_IMG_DIR = Path(r"W:\py\25.11_QeQ2\web2\book_img")

# C,E,G,... (id1~id10) 에 해당하는 DB 이미지들이 모여 있는 폴더
CANDIDATE_IMG_DIR = Path(r"W:\py\25.11_QeQ2\web2\candi")

# 샘플 이미지 파일명 패턴
#   - 엑셀에 "012p-08.png" 이런 식으로 확장자까지 들어 있으면: "{name}"
#   - 엑셀에 "012p-08" 까지만 있으면: "{name}.png"
SAMPLE_NAME_PATTERN = "{name}.png"        # 필요하면 "{name}.png" 로

# 후보(DB) 이미지 파일명 패턴
#   - 예: A3101B0013_q.png  이라면 "{name}_q.png"
#   - 예: A3101B0013.png    이라면 "{name}.png"
CANDIDATE_NAME_PATTERN = "{name}_p.gif"

# 출력 HTML
OUTPUT_HTML = Path(r"W:\py\25.11_QeQ2\web2\지수test_top10_gallery.html")





# ==========================
# 2. 유틸
# ==========================

def _rel_path_for_html(target_path: Path, html_path) -> str:
    html_path = Path(html_path)  # 문자열이 와도 Path로 바꿔줌
    try:
        rel = os.path.relpath(target_path, start=html_path.parent)
    except ValueError:
        rel = str(target_path)
    return rel.replace("\\", "/")



def _build_sample_img_path(name: str) -> Path:
    # name: 엑셀 A열 값
    name = name.strip()
    filename = SAMPLE_NAME_PATTERN.format(name=name)
    p = SAMPLE_IMG_DIR / filename
    return p


def _build_candidate_img_path(name: str) -> Path:
    # name: id1, id2, ... (A3101B0013 이런 것)
    name = name.strip()
    filename = CANDIDATE_NAME_PATTERN.format(name=name)
    p = CANDIDATE_IMG_DIR / filename
    return p


# ==========================
# 3. 메인
# ==========================

def main():
    print("[STEP] 엑셀 로딩:", EXCEL_PATH)
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    # 필요한 컬럼 체크
    if SAMPLE_COL not in df.columns:
        raise ValueError(f"엑셀에 '{SAMPLE_COL}' 열이 없습니다.")

    for c in SCORE_COLS + ID_COLS:
        if c not in df.columns:
            print(f"[WARN] 엑셀에 '{c}' 열이 없습니다. -> 비어 있는 값으로 처리")
            df[c] = None

    rows = []

    print("[STEP] 행 파싱")
    for _, row in df.iterrows():
        sample_name = row[SAMPLE_COL]
        if pd.isna(sample_name) or str(sample_name).strip() == "":
            continue

        sample_name = str(sample_name).strip()
        sample_path = _build_sample_img_path(sample_name)

        if not sample_path.exists():
            print(f"[WARN] 샘플 이미지 없음: {sample_path}")
            # 그래도 행은 남기고, 나중에 빈 셀로 표시
        cand_infos = []
        for score_col, id_col in zip(SCORE_COLS, ID_COLS):
            score_val = row[score_col]
            id_val = row[id_col]

            if pd.isna(id_val) or str(id_val).strip() == "":
                cand_infos.append((None, None, None))
                continue

            id_str = str(id_val).strip()
            score = None if pd.isna(score_val) else float(score_val)
            img_path = _build_candidate_img_path(id_str)
            if not img_path.exists():
                print(f"[WARN] 후보 이미지 없음: {img_path}")
            cand_infos.append((score, id_str, img_path))

        rows.append((sample_name, sample_path, cand_infos))

    if not rows:
        print("[ERROR] 유효한 행이 없습니다.")
        return

    print(f"[INFO] 유효한 행 개수: {len(rows)}")

    # ==========================
    # HTML 생성
    # ==========================

    print("[STEP] HTML 생성")
    html = []

    html.append("<!DOCTYPE html>")
    html.append("<html lang='ko'>")
    html.append("<head>")
    html.append("<meta charset='UTF-8'>")
    html.append("<title>지수test Top10 Gallery</title>")
    html.append("<style>")
    html.append("""
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: 0;
  font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  background-color: #111;
  color: #eee;
}
.table-container {
  width: 100vw;
  height: 100vh;
  overflow: auto;  /* 가로/세로 스크롤 둘 다 */
  padding: 8px;
}
table {
  border-collapse: collapse;
  background-color: #111;
  min-width: 1600px;  /* 가로 스크롤 유도 */
}
th, td {
  border: 1px solid #333;
  padding: 4px;
  text-align: center;
  vertical-align: top;
}
th {
  position: sticky;
  top: 0;
  background-color: #222;
  z-index: 10;
}
td {
  background-color: #181818;
}
img {
  display: block;
  max-width: 260px;
  max-height: 260px;
  width: auto;
  height: auto;
}
.caption {
  font-size: 12px;
  color: #ccc;
  margin-top: 4px;
  word-break: break-all;
}
.sample-header {
  font-weight: 600;
}
    """)
    html.append("</style>")
    html.append("</head>")
    html.append("<body>")
    html.append("<div class='table-container'>")
    html.append("<table>")

    # 헤더
    html.append("<thead>")
    html.append("<tr>")
    html.append("<th>원본</th>")
    for i in range(1, 11):
        html.append(f"<th>Top{i}</th>")
    html.append("</tr>")
    html.append("</thead>")

    # 바디
    html.append("<tbody>")
    for sample_name, sample_path, cand_infos in rows:
        html.append("<tr>")

        # 원본 셀
        if sample_path.exists():
            src = _rel_path_for_html(sample_path, OUTPUT_HTML)
            html.append("<td>")
            html.append(f"<div class='sample-header'>{sample_name}</div>")
            html.append(f"<img src='{src}' loading='lazy'>")
            html.append("</td>")
        else:
            html.append(f"<td><div class='sample-header'>{sample_name}</div><div>(이미지 없음)</div></td>")

        # 후보 10개
        for score, cand_id, img_path in cand_infos:
            if img_path is not None and img_path.exists():
                src = _rel_path_for_html(img_path, OUTPUT_HTML)
                html.append("<td>")
                html.append(f"<img src='{src}' loading='lazy'>")
                cap = cand_id
                if score is not None:
                    cap = f"{cand_id}<br>score={score:.4f}"
                html.append(f"<div class='caption'>{cap}</div>")
                html.append("</td>")
            else:
                # 이미지 없으면 캡션만 혹은 빈 셀
                if cand_id is None:
                    html.append("<td></td>")
                else:
                    cap = cand_id
                    if score is not None:
                        cap = f"{cand_id}<br>score={score:.4f}"
                    html.append(f"<td><div class='caption'>{cap}<br>(이미지 없음)</div></td>")

        html.append("</tr>")
    html.append("</tbody>")

    html.append("</table>")
    html.append("</div>")
    html.append("</body>")
    html.append("</html>")

    print("[STEP] HTML 저장:", OUTPUT_HTML)
    OUTPUT_HTML.write_text("\n".join(html), encoding="utf-8")

    print("[DONE] 지수test Top10 갤러리 생성 완료")


if __name__ == "__main__":
    main()
