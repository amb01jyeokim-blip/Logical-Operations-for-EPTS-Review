# -*- coding: utf-8 -*-
"""
엑셀 불리언 검색 FILTER 함수식 생성기 (Streamlit 웹앱)

실행 방법
    pip install -U streamlit
    streamlit run excel_boolean_filter.py

검색어 규칙 (한 칸 안에서)
    *   : AND  (예: 블록체인*보안)
    +   : OR   (예: AI+인공지능)
    ~   : NOT  (예: ~빅데이터)
    ( ) : 묶음 (예: (AI+인공지능)*~빅데이터)
    연산순서 ~ > * > + , 조건 간 AND 연결

날짜 조건
    날짜 범위와 시작일·종료일을 고르면 ((범위)>=DATE(연,월,일)) 형태의 조건으로 들어간다.
    시작일만 또는 종료일만 골라도 되고, 둘 다 비우면 조건에서 빠진다.

출력되는 함수식
    값 고정형 : 이 화면에 적은 검색어를 함수식 안에 그대로 넣는다.
    셀 참조형 : 엑셀 셀(검색칸)에 적은 검색어를 엑셀이 직접 해석한다. 검색어를 바꿔도 함수식은 그대로 둔다.

경로 입력
    엑셀 범위 (예: A2:A100, data!A2:A100)
    함수식에는 절대참조로 바뀌어 들어간다 (예: data!$A$2:$A$100)
    쉼표로 여러 범위를 적으면 이어 붙여서 검색 (예: A2:A100, C2:C100)
"""
import re
from datetime import date

import streamlit as st

OPERATORS = "+*~()"

STATIC = "값 고정형 (검색어 내장)"
DYNAMIC = "셀 참조형 (엑셀 검색칸 연동)"
MULTILINE = "줄바꿈"
ONELINE = "한 줄"

# 셀 참조형에서 검색칸의 문자열을 해석하는 재귀 LAMBDA.
# 괄호 밖의 + → 괄호 밖의 * → 맨 앞의 ~ → 바깥 괄호 순으로 쪼개고, 남은 단어를 SEARCH 한다.
# (LET 이름으로 r, c 는 쓸 수 없어 세 글자 이름을 사용한다.)
PARSER_LINES = [
    "LAMBDA(self,qry,rng,LET(",
    "txt,TRIM(qry),",
    "cnt,LEN(txt),",
    "chr,MID(txt,SEQUENCE(MAX(cnt,1)),1),",
    'lvl,SCAN(0,chr,LAMBDA(acc,x,acc+(x="(")-(x=")"))),',
    'pOr,IFERROR(MATCH(1,(chr="+")*(lvl=0),0),0),',
    'pAnd,IFERROR(MATCH(1,(chr="*")*(lvl=0),0),0),',
    "IF(pOr,SIGN(self(self,LEFT(txt,pOr-1),rng)+self(self,MID(txt,pOr+1,cnt),rng)),",
    "IF(pAnd,self(self,LEFT(txt,pAnd-1),rng)*self(self,MID(txt,pAnd+1,cnt),rng),",
    'IF(LEFT(txt)="~",1-self(self,MID(txt,2,cnt),rng),',
    'IF((LEFT(txt)="(")*(RIGHT(txt)=")"),self(self,MID(txt,2,cnt-2),rng),',
    "--ISNUMBER(SEARCH(txt,rng))))))))",
]

# 셀(A2) / 열(A) / 행(2) 주소. 이미 $가 붙어 있어도 인식한다.
ADDRESS = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)|\$?([A-Za-z]{1,3})|\$?(\d+)")


# ─────────────────────────── 검색어 해석 ───────────────────────────
class QueryError(ValueError):
    """검색어 문법 오류."""


def tokenize(text: str) -> list[str]:
    """검색어를 연산자와 단어로 나눈다."""
    tokens, word = [], ""
    for ch in text:
        if ch in OPERATORS:
            if word.strip():
                tokens.append(word.strip())
            tokens.append(ch)
            word = ""
        else:
            word += ch
    if word.strip():
        tokens.append(word.strip())
    return tokens


def join(kind: str, items: list) -> tuple:
    """같은 연산이 겹치면 한 단계로 펼친다. 예: (A+B)+C → A+B+C"""
    if len(items) == 1:
        return items[0]
    flat = []
    for item in items:
        flat += item[1] if item[0] == kind else [item]
    return (kind, flat)


def parse(text: str) -> tuple:
    """검색어를 트리로 변환한다. 노드: ("term", 단어) / ("not", 노드) / ("and"|"or", [노드...])"""
    tokens = tokenize(text)
    pos = 0

    def peek():
        return tokens[pos] if pos < len(tokens) else None

    def take() -> str:
        nonlocal pos
        pos += 1
        return tokens[pos - 1]

    def chain(operator: str, kind: str, operand) -> tuple:
        items = [operand()]
        while peek() == operator:
            take()
            items.append(operand())
        return join(kind, items)

    def expr() -> tuple:
        return chain("+", "or", term)

    def term() -> tuple:
        return chain("*", "and", factor)

    def factor() -> tuple:
        token = peek()
        if token is None:
            raise QueryError("연산자 뒤에 검색어가 없습니다.")
        if token in ")+*":
            raise QueryError(f"'{token}' 앞에 검색어가 없습니다.")
        take()
        if token == "~":
            return ("not", factor())
        if token == "(":
            node = expr()
            if peek() != ")":
                raise QueryError("닫는 괄호 ')'가 없습니다.")
            take()
            return node
        return ("term", token)

    node = expr()
    if peek() == ")":
        raise QueryError("여는 괄호 '('가 없습니다.")
    if peek() is not None:
        raise QueryError(f"'{peek()}' 앞에 연산자(*, +)가 없습니다.")
    return node


# ─────────────────────────── 함수식 생성 로직 ───────────────────────────
def split_csv(text: str) -> list[str]:
    """쉼표로 구분된 문자열을 리스트로 변환한다."""
    return [t.strip() for t in (text or "").replace("，", ",").split(",") if t.strip()]


def quote(text: str) -> str:
    """엑셀 문자열 리터럴로 변환한다 (큰따옴표 이스케이프)."""
    return '"' + text.replace('"', '""') + '"'


def absolute(ref: str) -> str:
    """범위를 절대참조로 바꾼다. 예: data!F2:F100 → data!$F$2:$F$100, A:A → $A:$A, E1 → $E$1
    셀 주소가 아닌 입력(이름 정의, 표 참조, 함수식 등)은 그대로 둔다."""
    ref = ref.strip()
    sheet, bang, address = ref.rpartition("!")
    parts = [ADDRESS.fullmatch(p.strip()) for p in address.split(":")]
    if len(parts) > 2 or not all(parts):
        return ref
    kinds = {"cell" if m.group(1) else "col" if m.group(3) else "row" for m in parts}
    if len(kinds) > 1 or (len(parts) == 1 and kinds != {"cell"}):
        return ref

    def fix(m: re.Match) -> str:
        if m.group(1):
            return f"${m.group(1).upper()}${m.group(2)}"
        if m.group(3):
            return f"${m.group(3).upper()}"
        return f"${m.group(4)}"

    return sheet + bang + ":".join(fix(m) for m in parts)


def range_expr(path: str) -> str:
    """경로 입력을 절대참조 범위식으로 변환한다. 여러 범위는 공백을 두고 이어 붙인다."""
    return '&" "&'.join(absolute(p) for p in split_csv(path))


def date_expr(day: date) -> str:
    """날짜를 엑셀 DATE 함수로 변환한다. 예: 2026-06-09 → DATE(2026,6,9)"""
    return f"DATE({day.year},{day.month},{day.day})"


def node_cond(node: tuple, rng: str, gap: str = "") -> str:
    """트리를 엑셀 조건식으로 변환한다. gap은 쉼표·연산자 사이에 넣는 공백이다."""
    kind, body = node
    if kind == "term":
        return f"ISNUMBER(SEARCH({quote(body)},{gap}{rng}))"
    if kind == "not":
        return f"NOT({node_cond(body, rng, gap)})"
    parts = [node_cond(n, rng, gap) for n in body]
    if kind == "and":
        return "(" + f"{gap}*{gap}".join(parts) + ")"
    return "((" + f"{gap}+{gap}".join(parts) + f"){gap}>{gap}0)"


def field_conds(node: tuple, rng: str, multiline: bool) -> list[str]:
    """한 검색칸의 조건식 목록을 만든다.
    줄바꿈 표시에서는 바깥 AND를 조건별로 나누고, 바깥 OR은 검색어마다 줄을 바꾼다."""
    if not multiline:
        return [node_cond(node, rng)]
    conds = []
    for n in (node[1] if node[0] == "and" else [node]):
        if n[0] == "or":
            parts = [node_cond(x, rng, " ") for x in n[1]]
            conds.append("((" + " +\n      ".join(parts) + ") > 0)")
        else:
            conds.append(node_cond(n, rng, " "))
    return conds


def date_conds(v: dict) -> tuple[list[str], list[str]]:
    """날짜 조건을 (조건식 목록, 오류목록)으로 변환한다. 예: ((data!$K$2:$K$100)>=DATE(2026,6,9))"""
    conds, errors = [], []
    for d in v["dates"]:
        if not (d["start"] or d["end"]):
            continue
        ranges = split_csv(d["path"])
        if not ranges:
            errors.append(f"{d['label']}의 날짜 범위가 비어 있습니다.")
        elif len(ranges) > 1:
            errors.append(f"{d['label']}의 날짜 범위는 하나만 입력하세요.")
        elif d["start"] and d["end"] and d["start"] > d["end"]:
            errors.append(f"{d['label']}의 시작일이 종료일보다 늦습니다.")
        else:
            rng = absolute(ranges[0])
            if d["start"]:
                conds.append(f"(({rng})>={date_expr(d['start'])})")
            if d["end"]:
                conds.append(f"(({rng})<={date_expr(d['end'])})")
    return conds, errors


def node_text(node: tuple, top: bool = True) -> str:
    """트리를 AND / OR / NOT 표기로 변환한다."""
    kind, body = node
    if kind == "term":
        return body
    if kind == "not":
        return "NOT " + node_text(body, top=False)
    text = f" {kind.upper()} ".join(node_text(n, top=False) for n in body)
    return text if top else f"({text})"


def filter_formula(ret: str, conds: list[str], empty_msg: str, multiline: bool,
                   indent: str = "") -> str:
    """조건식들을 AND로 묶어 FILTER 식을 만든다."""
    if not multiline:
        return f"FILTER({ret}," + "*".join(conds) + f",{quote(empty_msg)})"
    pad = indent + "    "
    return (f"FILTER({ret},\n{pad}" + f" *\n{pad}".join(conds)
            + f",\n{indent}{quote(empty_msg)})")


def check_common(v: dict, active: list[dict]) -> list[str]:
    errors = []
    if not v["ret_path"].strip():
        errors.append("결과 반환 범위가 비어 있습니다.")
    for f in active:
        if not range_expr(f["path"]):
            errors.append(f"{f['label']}의 검색 대상 범위가 비어 있습니다.")
    return errors


def build_preview(v: dict) -> tuple[str, list[str]]:
    """입력된 검색어와 날짜 조건을 (Query Preview, 오류목록)으로 변환한다."""
    lines, errors = [], []
    for f in v["fields"]:
        if not f["query"].strip():
            continue
        try:
            lines.append(f"({f['label']}: {node_text(parse(f['query']))})")
        except QueryError as e:
            errors.append(f"{f['label']}: {e}")
    for d in v["dates"]:
        if d["start"]:
            lines.append(f"({d['label']} >= {d['start'].isoformat()})")
        if d["end"]:
            lines.append(f"({d['label']} <= {d['end'].isoformat()})")
    return "\nAND ".join(lines), errors


def build_dynamic(v: dict) -> tuple[str, list[str]]:
    """셀 참조형: 검색칸 셀의 검색어를 엑셀이 해석하는 함수식. 날짜 조건은 DATE로 넣는다."""
    active = [f for f in v["fields"] if f["cell"].strip()]
    errors = check_common(v, active)
    if not active:
        errors.append("검색칸 셀을 하나 이상 입력하세요.")
    dconds, derrors = date_conds(v)
    errors += derrors
    if errors:
        return "", errors

    conds = [f"BSEARCH(BSEARCH,{absolute(f['cell'])},{range_expr(f['path'])})" for f in active]
    conds += dconds
    ret = absolute(v["ret_path"])
    if v["multiline"]:
        parser = "\n    ".join(PARSER_LINES)
        body = filter_formula(ret, conds, v["empty_msg"], True, indent="  ")
        return f"=LET(\n  BSEARCH,{parser},\n  {body})", []
    body = filter_formula(ret, conds, v["empty_msg"], False)
    return "=LET(BSEARCH," + "".join(PARSER_LINES) + f",{body})", []


def build_static(v: dict) -> tuple[str, list[str]]:
    """값 고정형: 입력한 검색어와 날짜를 그대로 넣은 함수식."""
    active = [f for f in v["fields"] if f["query"].strip()]
    errors = check_common(v, active)
    conds = []
    for f in active:
        try:
            conds += field_conds(parse(f["query"]), range_expr(f["path"]), v["multiline"])
        except QueryError as e:
            errors.append(f"{f['label']}: {e}")
    dconds, derrors = date_conds(v)
    conds += dconds
    errors += derrors
    if not active and not dconds and not derrors:
        errors.append("검색어 또는 날짜 조건을 하나 이상 입력하세요.")
    if errors:
        return "", errors
    return "=" + filter_formula(absolute(v["ret_path"]), conds, v["empty_msg"], v["multiline"]), []


# ─────────────────────────────── 화면 ───────────────────────────────
DEFAULT_FIELDS = [
    # (항목명, 검색어, 검색 대상 범위, 검색칸 셀)
    ("발간처", "기획예산처+금융위원회",
     "data!F2:F20196", "B1"),
    ("초록", "(AI*인공지능)+중소기업*~교육", "data!I2:I20196", "B2"),
]

DEFAULT_DATES = [
    # (항목명, 날짜 범위, 시작일, 종료일)
    ("발간일", "data!K2:K20196", date(2026, 6, 9), None),
]

RULES = """
| 기호 | 의미 | 예 |
|:---:|:---:|:---|
| `*` | AND | `블록체인*보안` |
| `+` | OR | `AI+인공지능` |
| `~` | NOT | `~빅데이터` |
| `( )` | 묶음 | `(AI+인공지능)*~빅데이터` |

우선순위는 `~` > `*` > `+` 이고, 칸과 칸은 AND로 연결됩니다.
"""


def field(i: int, dynamic: bool) -> None:
    """[필드명 | 키워드 | 필드 범위 | 검색칸 셀] 한 줄을 그린다."""
    if i < len(DEFAULT_FIELDS):
        label, query, path, cell = DEFAULT_FIELDS[i]
    else:
        label, query, path, cell = f"검색칸{i + 1}", "", "", ""
    c1, c2, c3, c4 = st.columns([1, 3, 1.5, 1])
    c1.text_input("컬럼명", value=label, key=f"label{i}")
    c2.text_input("검색어", value=query, key=f"query{i}", placeholder="예: (AI+인공지능)*~빅데이터")
    c3.text_input("범위", value=path, key=f"path{i}", placeholder="예: data!F2:F20196")
    c4.text_input("검색셀(참조형 입력)", value=cell, key=f"cell{i}", placeholder="예: B1", disabled=not dynamic)


def date_field(i: int) -> None:
    """[필드명 | 날짜 범위 | 시작일 | 종료일] 한 줄을 그린다."""
    if i < len(DEFAULT_DATES):
        label, path, start, end = DEFAULT_DATES[i]
    else:
        label, path, start, end = f"날짜{i + 1}", "", None, None
    limits = dict(min_value=date(1900, 1, 1), max_value=date(2100, 12, 31), format="YYYY-MM-DD")
    c1, c2, c3, c4 = st.columns([1, 3, 1.5, 1.5])
    c1.text_input("컬럼명", value=label, key=f"dlabel{i}")
    c2.text_input("발간일 범위", value=path, key=f"dpath{i}", placeholder="예: data!K2:K20196")
    c3.date_input("시작일", value=start, key=f"dstart{i}", **limits)
    c4.date_input("종료일", value=end, key=f"dend{i}", **limits)


def sidebar() -> None:
    with st.sidebar:
        st.subheader("Setting")
        st.text_input("메타데이터 범위", value="data!A2:AL20196",
                      key="ret_path")
        st.text_input("결과 없음 문구", value="검색 결과가 없습니다.", key="empty_msg")
        st.selectbox("함수식 표시", [MULTILINE, ONELINE], key="layout")
        st.divider()
        st.subheader("검색어 규칙")
        with st.expander("내부연산자 안내"):
            st.markdown(RULES)


def show(formula: str, errors: list[str]) -> None:
    if errors:
        for e in errors:
            st.error(e)
        return
    st.code(formula, language="excel")
    if len(formula) > 8192:
        st.warning(f"함수식 길이({len(formula):,}자)가 엑셀 한도(8,192자)를 넘습니다.")


def main() -> None:
    st.set_page_config(page_title="Query Builder for EPTS", page_icon="🔎", layout="wide")
    sidebar()

    st.title("🔎 Boolean Query Builder")
    st.caption("불리언 연산자를 활용해 체계적으로 정부 보도자료를 선별하는 함수식 생성도구입니다.")
    mode = st.radio("함수식 유형", [STATIC, DYNAMIC], horizontal=True, key="mode")
    dynamic = mode == DYNAMIC

    st.subheader("검색조건")
    count = st.number_input("검색조건 개수", min_value=1, max_value=10, value=2, key="count")
    for i in range(count):
        field(i, dynamic)

    st.subheader("발간일 조건")
    date_count = st.number_input("발간일 조건 개수", min_value=0, max_value=5, value=1, key="date_count")
    for i in range(date_count):
        date_field(i)

    values = {k: st.session_state[k] for k in ("ret_path", "empty_msg")}
    values["multiline"] = st.session_state["layout"] == MULTILINE
    values["fields"] = [
        {k: st.session_state[f"{k}{i}"] for k in ("label", "query", "path", "cell")}
        for i in range(count)
    ]
    for i, f in enumerate(values["fields"]):
        f["label"] = f["label"].strip() or f"검색칸{i + 1}"
    values["dates"] = [
        {k: st.session_state[f"d{k}{i}"] for k in ("label", "path", "start", "end")}
        for i in range(date_count)
    ]
    for i, d in enumerate(values["dates"]):
        d["label"] = d["label"].strip() or f"날짜{i + 1}"

    preview, preview_errors = build_preview(values)
    st.subheader("연산식 Query Preview")
    st.code(preview or " ", language=None)
    for e in preview_errors:
        st.error(e)

    st.subheader("함수식 Search Filter")
    if dynamic:
        st.info("검색셀에 검색어를 적으면 결과가 바로 연동됩니다. 빈 검색칸은 조건에서 빠집니다. "
                "Microsoft 365 또는 Excel 2024 (LAMBDA 지원 버전)에서 동작합니다.")
        show(*build_dynamic(values))
    else:
        st.info("위에 입력한 검색어가 함수식에 그대로 들어갑니다. "
                "FILTER를 지원하는 Excel 2021 이상에서 동작합니다.")
        show(*build_static(values))


if __name__ == "__main__":
    main()

#2026.11월 KDI 인턴십