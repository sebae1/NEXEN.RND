import json
from pydantic import BaseModel
from openai import OpenAI
from anthropic import Anthropic

KEYMAP = {
    0: "Cost Category PK",
    1: "Cost Category Name",
    2: "Cost Category Path",
    3: "Cost Element PK",
    4: "Cost Element Description",
    5: "Cost Ctr PK",
    6: "Cost Ctr Name",
    7: "Cost Ctr RND Type",
    8: "Cost Ctr OE Type",
    9: "Cost Ctr Path",
    10: "RND Type PK",
    11: "RND Type Name",
    12: "OE Type PK",
    13: "OE Type Name",
    14: "Month",
    15: "계획금액",
    16: "집행금액"
}

class _CostCategory(BaseModel):
    pk: int
    parent_pk: int|None
    name: str

class _CostElement(BaseModel):
    code: str
    description: str
    category_pk: int|None

class _CostCtr(BaseModel):
    code: str
    parent_code: str|None
    name: str
    rnd: int
    oe: int

class _BudgetByElement(BaseModel):
    cost_element_code: str
    planned: float
    executed: float

class _BudgetByCtr(BaseModel):
    cost_ctr_code: str
    planned: float
    executed: float

def get_prompts_for_ai(
       cost_category_list: list[_CostCategory],
       cost_element_list: list[_CostElement],
       cost_ctr_list: list[_CostCtr],
       budget_by_element: list[_BudgetByElement],
       budget_by_ctr: list[_BudgetByCtr],
    ) -> tuple[str, str, str]:
    """
    Returns:
        system_prompt
        user_prompt
        json_data
    """
    category_map = {cat.pk: cat for cat in cost_category_list}
    def get_category_path(leaf_pk: int) -> str:
        cat = category_map[leaf_pk]
        path = [str(cat.pk),]
        while True:
            if cat.parent_pk is None:
                return ".".join(path)
            cat = category_map[cat.parent_pk]
            path.insert(0, str(cat.pk))

    elem_code_vs_pk = {elem.code: idx for idx, elem in enumerate(cost_element_list)}

    ctr_map: dict[int, _CostCtr] = {}
    ctr_code_vs_pk: dict[str, int] = {}
    for idx, ctr in enumerate(cost_ctr_list):
        ctr_map[idx] = ctr
        ctr_code_vs_pk[ctr.code] = idx
    def get_ctr_path(leaf_code: str) -> str:
        pk = ctr_code_vs_pk[leaf_code]
        ctr = ctr_map[pk]
        path = [str(pk),]
        while True:
            if ctr.parent_code is None:
                return ".".join(path)
            pk = ctr_code_vs_pk[ctr.parent_code]
            ctr = ctr_map[pk]
            path.insert(0, str(pk))

    json_data = {
        "keymap": KEYMAP,
        "cost_category": [
            {
                0: cat.pk,
                1: cat.name,
            } for cat in cost_category_list
        ],
        "cost_element": [
            {
                2: get_category_path(elem.category_pk),
                3: elem_code_vs_pk[elem.code],
                4: elem.description
            } for elem in cost_element_list if elem.category_pk in category_map
        ],
        "cost_ctr": [
            {
                5: ctr_code_vs_pk[ctr.code],
                6: ctr.name,
                7: ctr.rnd,
                8: ctr.oe
            } for ctr in cost_ctr_list
        ],
        "rnd_type": [
            {10: 0, 11: "Research"},
            {10: 1, 11: "Develop"},
        ],
        "oe_type": [
            {12: 0, 13: "공통비"},
            {12: 1, 13: "RE"},
            {12: 2, 13: "OE"},
        ],
        "budget_by_element": [
            {
                3: elem_code_vs_pk[item.cost_element_code],
                15: item.planned,
                16: item.executed
            } for item in budget_by_element if item.cost_element_code in elem_code_vs_pk
        ],
        "budget_by_ctr": [
            {
                9: get_ctr_path(item.cost_ctr_code),
                15: item.planned,
                16: item.executed
            } for item in budget_by_ctr if item.cost_ctr_code in ctr_code_vs_pk
        ],
    }
    compact_json = json.dumps(json_data, ensure_ascii=False, separators=(",", ":"))

    system_prompt = """\
당신은 재무/예산 분석가입니다.
아래 JSON의 숫자 키는 'keymap'에 정의된 의미로 해석해야 합니다.
- 'cost_category': {0: 카테고리PK, 1: 카테고리명}. 카테고리 경로(2)는 cost_element 또는 budget_by_element에서 제공됩니다.
- 'cost_element': {2: 카테고리PK경로, 3: 원가요소PK, 4: 설명}.
- 'cost_ctr': {5: 센터PK, 6: 센터이름, 7: RND타입PK, 8: OE타입PK}. 센터 경로는 budget_item에서 9번으로 제공됩니다.
- 'rnd_type': {10: RND타입PK, 11: 이름}, 'oe_type': {12: OE타입PK, 13: 이름}.
- 'budget_by_element': {3: 원가요소PK, 15: 계획, 16: 집행}.
- 'budget_by_ctr': {9: 센터PK경로, 15: 계획, 16: 집행}.
숫자 연산/집계는 정확히 수행하되, 불확실한 가정은 하지 마세요.
PK 요약 내용에 사용하지 말고 항상 이름(Name)을 사용하세요.
출력은 **마크다운**으로만 작성합니다.
가능하다면 표를 적극적으로 포함합니다."""

    user_prompt = f"""\
아래 JSON 데이터를 분석하여 '계획 vs 집행' 재무 분석 보고서를 작성하세요.
필수 섹션:
1) 요약
2) 센터 관점 요약
3) 원가요소 관점 요약
4) 리스크 & 권고안
5) 부록

데이터(JSON):
{json_data}"""

    return system_prompt, user_prompt, compact_json

def get_gpt_models(key: str) -> list[str]:
    client = OpenAI(api_key=key)
    models = client.models.list()
    model_names = [m.id for m in models.data if m.id.startswith("gpt")]
    return model_names

def analyze_by_gpt(system_prompt: str, user_prompt: str, key: str, model: str) -> str:
    client = OpenAI(api_key=key)

    resp = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        max_output_tokens=2000,
    )

    markdown = resp.output_text
    return markdown

def get_claude_models(key: str) -> list[str]:
    client = Anthropic(api_key=key)
    models = client.models.list()
    model_names = [m.id for m in models]
    return model_names

def analyze_by_claude(system_prompt: str, user_prompt: str, key: str, model: str) -> str:
    """
    Claude API를 이용해 GPT 버전과 동일한 방식으로 분석 결과를 반환.
    - system_prompt : Claude의 역할 지시
    - user_prompt   : 실제 분석 요청
    - key           : Anthropic API Key
    - model         : 예) "claude-3-sonnet-20240229", "claude-3-opus-20240229"
    """
    client = Anthropic(api_key=key)

    resp = client.messages.create(
        model=model,
        max_tokens=2000,
        system=system_prompt,
        messages=[
            {"role": "user", "content": user_prompt}
        ]
    )

    # Claude의 응답 본문 추출
    content_blocks = resp.content
    assert content_blocks, "No content"

    return content_blocks[0].text

    # first = content_blocks[0]
    # if isinstance(first, dict):
    #     return first["text"]
    # else:
    #     # 객체 블록 (예: TextBlock)
    #     return first.text
