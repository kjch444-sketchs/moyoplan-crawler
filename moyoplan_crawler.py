import time
import re
import pandas as pd

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


# ============================================================
# Selenium 설정
# ============================================================

options = Options()

# GitHub Actions / Linux 환경
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")

options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(options=options)


# ============================================================
# 데이터 저장
# ============================================================

all_plans_data = []
seen_plans = set()


# ============================================================
# 허용된 사업자 명단
# ============================================================

allowed_carriers = [
    "찬스모바일",
    "KT엠모바일",
    "쉐이크모바일",
    "LG헬로모바일",
    "핀다이렉트",
    "KT스카이라이프",
    "U+유모바일",
    "이야기모바일",
    "고고모바일",
    "KG모바일",
    "인스모바일",
    "이지모바일",
    "마블링",
    "모나",
    "티플러스",
    "슈가모바일",
    "더원모바일",
    "시월모바일",
    "스마텔",
    "위너스텔",
    "아이즈모바일",
    "에이모바일",
    "모빙",
    "리브모바일",
    "프리티"
]


# ============================================================
# 월 요금 추출
# ============================================================

def extract_monthly_price(lines, full_text):

    for line in lines:

        # 할인 종료 후 요금은 제외
        if "이후" in line or "부터" in line:
            continue

        match = re.search(
            r"월\s*([0-9,]+)\s*원",
            line
        )

        if match:

            price = match.group(1)

            price = re.sub(
                r"[^0-9]",
                "",
                price
            )

            if price:
                return price

    match = re.search(
        r"월\s*([0-9,]+)\s*원",
        full_text
    )

    if match:

        price = match.group(1)

        price = re.sub(
            r"[^0-9]",
            "",
            price
        )

        if price:
            return price

    return "정보 없음"


# ============================================================
# 데이터량 추출
# ============================================================

def extract_data_amount(lines, plan_name):

    if plan_name != "알 수 없음":

        try:

            plan_index = lines.index(plan_name)

            if plan_index + 1 < len(lines):

                next_line = lines[
                    plan_index + 1
                ].strip()

                data_pattern = re.search(
                    r"("
                    r"(?:월\s*)?"
                    r"\d[\d,.]*\s*(?:GB|MB|TB)"
                    r"(?:\s*\+\s*\d+\s*Mbps)?"
                    r"|"
                    r"(?:월\s*)?데이터\s*무제한"
                    r"|"
                    r"무제한"
                    r")",
                    next_line,
                    re.IGNORECASE
                )

                if data_pattern:
                    return next_line

        except ValueError:
            pass

    for line in lines:

        line = line.strip()

        if not line:
            continue

        if line == plan_name:
            continue

        if re.search(
            r"\d[\d,.]*\s*(GB|MB|TB)",
            line,
            re.IGNORECASE
        ):

            if "원" in line:
                continue

            if (
                "통화" in line
                or "문자" in line
                or "건" in line
            ):
                continue

            return line

        if "데이터 무제한" in line:
            return line

    return "정보 없음"


# ============================================================
# 할인 기간 추출
# ============================================================

def extract_after_months(lines):

    for line in lines:

        if "개월" not in line:
            continue

        if not (
            "이후" in line
            or "부터" in line
            or "변경" in line
        ):
            continue

        match = re.search(
            r"(\d+)\s*개월",
            line
        )

        if match:
            return match.group(1)

    return "정보 없음"


# ============================================================
# 기간 이후 요금 추출
# ============================================================

def extract_after_price(lines):

    for line in lines:

        if "원" not in line:
            continue

        if (
            "이후" not in line
            and "전환" not in line
        ):
            continue

        match = re.search(
            r"(?:이후|전환)\s*([0-9,]+)\s*원",
            line
        )

        if match:

            price = match.group(1)

            price = re.sub(
                r"[^0-9]",
                "",
                price
            )

            if price:
                return price

    return "정보 없음"


# ============================================================
# 페이백 정보 추출
# ============================================================

def extract_payback(lines, full_text):

    potential_paybacks = []

    for line in lines:

        if (
            "페이백" in line
            or "캐시백" in line
        ):
            potential_paybacks.append(line)

    if potential_paybacks:

        best_payback = potential_paybacks[0]

        for p in potential_paybacks:

            if len(p) > len(best_payback):
                best_payback = p

            elif (
                "원" in p
                or "포인트" in p
            ):
                best_payback = p

        return best_payback

    if "페이백" in full_text:

        match = re.search(
            r".{0,50}페이백.{0,100}",
            full_text
        )

        if match:
            return match.group(0).strip()

    if "캐시백" in full_text:

        match = re.search(
            r".{0,50}캐시백.{0,100}",
            full_text
        )

        if match:
            return match.group(0).strip()

    return "정보 없음"


# ============================================================
# 요금제 카드 여부 확인
# ============================================================

def is_plan_card(div):

    text = div.get_text(
        " ",
        strip=True
    )

    has_data = any(
        keyword in text
        for keyword in [
            "GB",
            "MB",
            "TB",
            "무제한"
        ]
    )

    has_call = any(
        keyword in text
        for keyword in [
            "통화",
            "분",
            "기본제공"
        ]
    )

    return has_data and has_call


# ============================================================
# 크롤링
# ============================================================

try:

    # 전체 페이지
    for page in range(220):

        url = (
            f"https://www.moyoplan.com/"
            f"plans?page.page={page}"
        )

        print()
        print("=" * 70)
        print(f"[{page}페이지] 접속 중...")
        print(url)
        print("=" * 70)

        driver.get(url)

        # 페이지 로딩
        time.sleep(4)

        soup = BeautifulSoup(
            driver.page_source,
            "html.parser"
        )

        # ====================================================
        # 요금제 카드 찾기
        # ====================================================

        plan_cards = []

        potential_cards = soup.find_all(
            "div",
            class_=True
        )

        for div in potential_cards:

            try:

                if not is_plan_card(div):
                    continue

                div_count = len(
                    div.find_all("div")
                )

                if div_count >= 30:
                    continue

                plan_cards.append(div)

            except Exception:
                continue

        # ====================================================
        # 중복 카드 제거
        # ====================================================

        unique_cards = []
        card_texts = set()

        for card in plan_cards:

            text = card.get_text(
                " ",
                strip=True
            )

            if text in card_texts:
                continue

            card_texts.add(text)
            unique_cards.append(card)

        plan_cards = unique_cards

        if not plan_cards:

            print(
                f"{page}페이지에서 요금제 카드를 "
                f"찾지 못했습니다."
            )

            continue

        print(
            f"[{page}페이지] 감지된 요금제 카드 수: "
            f"{len(plan_cards)}개"
        )

        page_count = 0

        # ====================================================
        # 카드별 데이터 추출
        # ====================================================

        for card in plan_cards:

            try:

                lines = [
                    line.strip()
                    for line in card.get_text(
                        separator="\n"
                    ).split("\n")
                    if line.strip()
                ]

                full_joined_text = " ".join(lines)

                carrier = "알 수 없음"
                plan_name = "알 수 없음"

                data = "정보 없음"
                call = "정보 없음"
                sms = "정보 없음"

                network = "알 수 없음"
                generation = "LTE"

                monthly_price = "정보 없음"

                after_months = "정보 없음"
                after_price_info = "정보 없음"

                payback_info = "정보 없음"

                # =================================================
                # 통신사 이미지 ALT
                # =================================================

                img_tag = card.find("img")

                if img_tag:

                    img_alt = img_tag.get("alt")

                    if img_alt:

                        img_alt = img_alt.strip()

                        if (
                            img_alt
                            and len(img_alt) <= 15
                            and not any(
                                x in img_alt
                                for x in [
                                    "점",
                                    "평점",
                                    "별점",
                                    "하트",
                                    "icon",
                                    "logo"
                                ]
                            )
                        ):
                            carrier = img_alt

                # =================================================
                # 별점 기준 요금제명 / 사업자명
                # =================================================

                for i, line in enumerate(lines):

                    is_rating = False

                    if "점" in line:
                        is_rating = True

                    elif (
                        line.replace(".", "").isdigit()
                        and len(line) <= 3
                    ):

                        try:

                            val = float(line)

                            if (
                                1.0
                                <= val
                                <= 5.0
                            ):
                                is_rating = True

                        except Exception:
                            pass

                    if is_rating:

                        if i + 1 < len(lines):

                            candidate_plan = (
                                lines[i + 1].strip()
                            )

                            if candidate_plan:
                                plan_name = candidate_plan

                        for j in range(
                            i - 1,
                            -1,
                            -1
                        ):

                            candidate = (
                                lines[j].strip()
                            )

                            if not candidate:
                                continue

                            if len(candidate) > 15:
                                continue

                            if "점" in candidate:
                                continue

                            if candidate.replace(
                                ".",
                                ""
                            ).isdigit():
                                continue

                            if any(
                                x in candidate
                                for x in [
                                    "GB",
                                    "MB",
                                    "TB"
                                ]
                            ):
                                continue

                            if "원" in candidate:
                                continue

                            carrier = candidate
                            break

                        break

                # =================================================
                # 보조 통신사 매칭
                # =================================================

                if (
                    carrier == "알 수 없음"
                    or not any(
                        ac in carrier
                        for ac in allowed_carriers
                    )
                ):

                    for line in lines:

                        found = False

                        for ac in allowed_carriers:

                            if ac in line:

                                carrier = ac
                                found = True
                                break

                        if found:
                            break

                # =================================================
                # 정확한 사업자명
                # =================================================

                matched_carrier = "알 수 없음"

                for ac in allowed_carriers:

                    if ac in carrier:

                        matched_carrier = ac
                        break

                if matched_carrier == "알 수 없음":
                    continue

                carrier = matched_carrier

                # =================================================
                # 요금제명 보정
                # =================================================

                if (
                    plan_name == "알 수 없음"
                    and len(lines) > 1
                ):
                    plan_name = lines[1]

                # =================================================
                # 데이터
                # =================================================

                data = extract_data_amount(
                    lines,
                    plan_name
                )

                # =================================================
                # LTE / 5G
                # =================================================

                if "5G" in full_joined_text:
                    generation = "5G"
                else:
                    generation = "LTE"

                # =================================================
                # 망 정보
                # =================================================

                if (
                    "SKT" in full_joined_text
                    or "SK망" in full_joined_text
                ):

                    network = "SKT"

                elif (
                    "KT" in full_joined_text
                    or "KT망" in full_joined_text
                ):

                    network = "KT"

                elif (
                    "LG" in full_joined_text
                    or "LGU+" in full_joined_text
                    or "LG망" in full_joined_text
                ):

                    network = "LG U+"

                # =================================================
                # 통화 / 문자
                # =================================================

                for line in lines:

                    if (
                        "통화" in line
                        or "분" in line
                        or "기본제공" in line
                    ):
                        call = line

                    if (
                        "문자" in line
                        or "건" in line
                    ):
                        sms = line

                # =================================================
                # 월 요금
                # =================================================

                monthly_price = (
                    extract_monthly_price(
                        lines,
                        full_joined_text
                    )
                )

                # =================================================
                # 할인 기간
                # =================================================

                after_months = (
                    extract_after_months(lines)
                )

                # =================================================
                # 기간 이후 요금
                # =================================================

                after_price = (
                    extract_after_price(lines)
                )

                if after_price != "정보 없음":

                    after_price_info = after_price

                # =================================================
                # 페이백
                # =================================================

                payback_info = (
                    extract_payback(
                        lines,
                        full_joined_text
                    )
                )

                # =================================================
                # 중복 제거
                # =================================================

                plan_key = (
                    carrier,
                    plan_name
                )

                if (
                    plan_key in seen_plans
                    or plan_name == "알 수 없음"
                ):
                    continue

                seen_plans.add(plan_key)

                # =================================================
                # 데이터 저장
                # =================================================

                all_plans_data.append({

                    "사업자 명":
                        carrier,

                    "요금제 명":
                        plan_name,

                    "통화제공량":
                        call,

                    "문자제공량":
                        sms,

                    "데이터제공량":
                        data,

                    "망정보":
                        network,

                    "LTE/5G 구분":
                        generation,

                    "월 요금":
                        monthly_price,

                    "할인 기간":
                        after_months,

                    "기간 이후 요금":
                        after_price_info,

                    "페이백 사은품":
                        payback_info
                })

                page_count += 1

                print(
                    f"  ✓ {carrier} | "
                    f"{plan_name} | "
                    f"데이터: {data} | "
                    f"월 {monthly_price}원"
                )

            except Exception as e:

                print(
                    f"  ⚠ 카드 처리 중 오류: {e}"
                )

                continue

        print()
        print(
            f"[{page}페이지] 파싱 완료 "
            f"→ 신규 {page_count}개"
        )

finally:

    driver.quit()


# ============================================================
# DataFrame
# ============================================================

df = pd.DataFrame(
    all_plans_data
)


# ============================================================
# 컬럼 순서
# ============================================================

columns = [
    "사업자 명",
    "요금제 명",
    "통화제공량",
    "문자제공량",
    "데이터제공량",
    "망정보",
    "LTE/5G 구분",
    "월 요금",
    "할인 기간",
    "기간 이후 요금",
    "페이백 사은품"
]

df = df[columns]


# ============================================================
# 월 요금 숫자형 변환
# ============================================================

df["월 요금"] = pd.to_numeric(
    df["월 요금"],
    errors="coerce"
)


# ============================================================
# 엑셀 저장
# ============================================================

output_file = "moyoplan_parsed_plans.xlsx"

df.to_excel(
    output_file,
    index=False
)


# ============================================================
# 결과
# ============================================================

print()
print("=" * 70)
print("크롤링 완료!")
print("=" * 70)

print(
    f"총 수집 요금제: {len(df)}개"
)

print(
    f"월 요금 수집 성공: "
    f"{df['월 요금'].notna().sum()}개"
)

print(
    f"월 요금 수집 실패: "
    f"{df['월 요금'].isna().sum()}개"
)

print(
    f"엑셀 저장 파일: {output_file}"
)

print("=" * 70)
