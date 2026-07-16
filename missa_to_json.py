import argparse
import json
import os
import re
import sys

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://missa.cbck.or.kr/DailyMissa/"

PSALM_SECTIONS = {"입당송", "화답송", "영성체송"}

TARGET_SECTIONS = [
    "입당송",
    "제1독서",
    "화답송",
    "제2독서",
    "복음 환호송",
    "복음",
    "영성체송",
]

OUTPUT_KEY_MAP = {
    "복음 환호송": "복음환호송",
}


def fetch_html(date_str: str) -> str:
    url = BASE_URL + date_str
    resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
    resp.raise_for_status()
    return resp.text


def truncate_psalm_title(title: str) -> str:
    """시편으로 시작하는 title은 첫 번째 콤마 앞까지만 반환."""
    if title.startswith("시편"):
        comma_idx = title.find(",")
        if comma_idx != -1:
            return title[:comma_idx].strip()
    return title


def extract_title_from_h4_span(h4) -> str:
    span = h4.find("span", class_="float-right")
    if span:
        return span.get_text(strip=True).replace("\xa0", " ").strip()
    return ""


def extract_content_lines(block) -> list[str]:
    tjustify = block.find("div", class_="tjustify")
    if not tjustify:
        return []
    lines = []
    # Only collect leaf divs (no child divs) to avoid duplicating parent text
    for div in tjustify.find_all("div"):
        if not div.find("div"):
            text = div.get_text(strip=True)
            if text:
                lines.append(text)
    return lines


def build_content(block) -> str:
    lines = extract_content_lines(block)
    return "\n".join(lines)


def find_section_block(soup, section_name: str):
    """section_name과 정확히 일치하는 h4를 찾아 상위 bottompadding-sm div 반환.

    h4의 직접 텍스트 노드(span 제외)만 비교하여 "복음"과 "복음 환호송"을 구분한다.
    HTML에 <div class="row bottompadding-sm">과 <div class="bottompadding-sm">이
    중첩되어 있으므로 class가 정확히 ['bottompadding-sm']인 div를 찾는다.
    """
    from bs4 import NavigableString
    for h4 in soup.find_all("h4"):
        direct_text = "".join(
            str(c) for c in h4.children if isinstance(c, NavigableString)
        ).strip()
        if direct_text == section_name:
            for parent in h4.parents:
                if parent.name == "div" and parent.get("class") == ["bottompadding-sm"]:
                    return h4, parent
    return None, None


def parse_reading_title(block) -> dict:
    """제1독서/제2독서/복음처럼 h4에 출처가 없는 섹션의 title 추출.

    h5.float-right 안의 구절 번호와, 그 h5를 포함하는 div의 책 이름 텍스트를 분리한다.
    예: "▥ 에제키엘 예언서의 말씀입니다." + "34,11-16"
        → {"title": "에제키엘 예언서의 말씀입니다.", "chapter_verse": "34,11-16"}
    예: "✠ 루카가 전한 거룩한 복음입니다." + "15,3-7"
        → {"title": "루카가 전한 거룩한 복음입니다.", "chapter_verse": "15,3-7"}
    """
    h5 = block.find("h5", class_="float-right")
    if not h5:
        return {"title": "", "chapter_verse": ""}
    chapter_verse = h5.get_text(strip=True)

    source_container = h5.find_parent("div")
    if source_container:
        full_text = source_container.get_text(strip=True)
        book_text = full_text.replace(chapter_verse, "").strip()
        # ▥/✠ 전례 기호 제거
        book_text = re.sub(r"^[▥✠]\s*", "", book_text).strip()
        return {"title": book_text, "chapter_verse": chapter_verse}

    return {"title": "", "chapter_verse": chapter_verse}


def parse_section(soup, section_name: str):
    h4, block = find_section_block(soup, section_name)
    if h4 is None or block is None:
        return None

    # Determine title
    title_from_h4 = extract_title_from_h4_span(h4)

    chapter_verse = None
    if title_from_h4:
        title = title_from_h4
        if section_name in PSALM_SECTIONS:
            title = truncate_psalm_title(title)
    else:
        # Readings and Gospel: use full "말씀입니다." text from HTML
        reading_info = parse_reading_title(block)
        title = reading_info["title"]
        chapter_verse = reading_info["chapter_verse"]

    content = build_content(block)

    if not content:
        return None

    if section_name in ("제1독서", "제2독서", "복음"):
        lines = content.split("\n")
        # 마지막 두 줄이 정해진 응답 쌍일 때만 제거 (중간에 나오는 경우는 제거하지 않음)
        if section_name in ("제1독서", "제2독서"):
            if (len(lines) >= 2
                    and "주님의 말씀입니다" in lines[-2]
                    and "하느님, 감사합니다" in lines[-1]):
                lines = lines[:-2]
        else:  # 복음
            if (len(lines) >= 2
                    and "주님의 말씀입니다" in lines[-2]
                    and "그리스도님, 찬미합니다" in lines[-1]):
                lines = lines[:-2]
        content = " ".join(lines)

    result = {"title": title, "content": content}
    if chapter_verse is not None:
        result["chapter_verse"] = chapter_verse
    return result


def extract_liturgy_name(soup) -> str:
    meta = soup.find("meta", attrs={"name": "title"})
    if meta and meta.get("content"):
        # "2026.06.14 [녹] 연중 제11주일" → "연중 제11주일"
        parts = meta["content"].split("]")
        if len(parts) > 1:
            return parts[-1].strip()
    return ""


def parse_missa(html: str, date_str: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")

    formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
    result = {
        "date": formatted_date,
        "liturgy": extract_liturgy_name(soup),
    }

    for section in TARGET_SECTIONS:
        key = OUTPUT_KEY_MAP.get(section, section)
        result[key] = parse_section(soup, section)

    return result


def save_json(data: dict, output_root: str, date_str: str) -> str:
    dir_path = os.path.join(output_root, date_str)
    os.makedirs(dir_path, exist_ok=True)
    file_path = os.path.join(dir_path, f"missa_{date_str}.json")
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return file_path


def run(date_str: str, output_root: str = '.') -> str:
    """매일미사 데이터를 가져와 JSON으로 저장. 저장된 파일 경로를 반환."""
    print(f"[{date_str}] 매일미사 데이터를 가져오는 중...")
    html = fetch_html(date_str)
    data = parse_missa(html, date_str)
    output_path = save_json(data, output_root, date_str)
    print(f"저장 완료: {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="매일미사 내용을 JSON으로 저장합니다.")
    parser.add_argument("date", help="날짜 (YYYYMMDD 형식, 예: 20260614)")
    parser.add_argument(
        "--output",
        default=".",
        help="출력 루트 디렉터리 (기본값: 현재 디렉터리). 날짜 서브디렉터리가 자동 생성됩니다.",
    )
    args = parser.parse_args()

    date_str = args.date
    if not re.fullmatch(r"\d{8}", date_str):
        print(f"오류: 날짜는 YYYYMMDD 형식이어야 합니다. (입력값: {date_str})", file=sys.stderr)
        sys.exit(1)

    try:
        run(date_str, args.output)
    except requests.HTTPError as e:
        print(f"오류: 페이지를 가져올 수 없습니다. ({e})", file=sys.stderr)
        sys.exit(1)
    except requests.RequestException as e:
        print(f"오류: 네트워크 오류가 발생했습니다. ({e})", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
