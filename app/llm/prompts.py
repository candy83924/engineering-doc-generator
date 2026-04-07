"""Prompt templates for LLM-powered content generation."""

# ============================================================
# ACCEPTANCE SPEC PROMPTS (5-section Word document)
# ============================================================

ACCEPTANCE_SYSTEM_PROMPT = """\
你是一位資深的台灣工程驗收規範撰寫專家。你的任務是根據報價單內容，撰寫一份專業、詳細、可直接用於正式工程驗收的規範文件。

你撰寫的內容必須：
1. 專業且具體，不可使用罐頭文字或空泛語句
2. 緊扣報價單中的實際工程內容、材料、品牌、規格、數量
3. 每一段描述都要包含報價單中的具體材料名稱、規格參數、施作數量
4. 若報價單資訊不完整，可在合理範圍內推定補足，但不可超出報價範圍

文件結構為五大區塊：
1. 維修設備基本資訊（表格：設備位置、設備類型、馬力規格、問題說明、維修需求）
2. 施工處置與安全規範（含四個子標題：施工廠商需具備、作業前準備、施工內容、驗收作業）
3. 施工與驗收標準表（表格：驗收項目 + 驗收標準說明）
4. 施工保險要求（一段完整敘述）
5. 施工區域（一段完整敘述）

回傳格式必須是純 JSON。
"""

ACCEPTANCE_USER_PROMPT_TEMPLATE = """\
請根據以下報價單內容，生成驗收規範所需的各項內容。

## 工程資訊
- 工程名稱：{project_name}
- 施工位置：{project_location}
- 廠商名稱：{vendor_name}
- 文件編號：{document_number}
- 案子概略說明：{project_brief}

## 報價項目明細
{items_detail}

## 附帶條件
{terms_detail}

## 要求輸出 JSON 格式

請嚴格按照以下 JSON 結構回傳：

{{
  "equipment_location": "（設備位置：具體描述施工範圍與位置，例如 'YM1 廠房棟 5 樓 L 區 1K to 100 級無塵室修改範圍，含裝修區、冰水管路區、FFU 配電區及中央監控區'）",
  "equipment_type": "（設備類型：列出本案涉及的所有設備與系統類型，例如 '無塵室裝修系統、庫板隔間、T-GRID、抗靜電高架地板、乾式冷卻盤管、冰水管路、FFU 配電盤與配線'）",
  "spec_info": "（馬力規格：概述本案設備規格依據，例如 '依現場既設設備及本次報價內容辦理；乾式冷卻盤管為 DCC1 規格，FFU、配電盤與監控點位依圖說及報價清單為準'）",
  "problem_description": "（問題說明：描述現場目前為何需要施工，現有設施有什麼不足、不符合需求的狀況，要具體描述，不可泛泛而談，至少 80 字）",
  "repair_needs": "（維修需求：描述這次要怎麼修改、處理、施作，要對應報價項目，具體列出主要工程項目，至少 80 字）",
  "contractor_requirements": [
    "（施工廠商需具備的資格、經驗與能力，每條都要具體。寫 4-6 條，每條至少 50 字，要包含具體工種、管理要求、人員資格等）"
  ],
  "preparation_steps": [
    "（作業前準備事項，每條都要包含具體材料名稱和動作。寫 6-8 條，每條至少 60 字）"
  ],
  "construction_content": [
    "（施工內容與步驟，這是最重要的部分。寫 10-15 條，每條至少 80 字。必須逐項對應報價單的工程項目，描述施作方式、使用材料、安裝位置、品質要求等）"
  ],
  "acceptance_procedures": [
    "（驗收作業要求，寫 5-7 條，每條至少 60 字。分別對應各工程系統的驗收重點）"
  ],
  "acceptance_criteria": [
    {{
      "category": "（驗收項目名稱，如：材料驗收、結構固定驗收、門框驗收、通用驗收、場地復原與清潔、工程保固）",
      "standard": "（驗收標準說明：以條列方式描述具體驗收標準，用換行分隔。每個類別寫 3-5 條標準，每條要具體可量化或可判定。格式為：1. xxx\\n2. xxx\\n3. xxx）"
    }}
  ],
  "insurance_text": "（施工保險要求：完整一段，描述投保要求、保險證明、安全管制等，至少 60 字）",
  "work_area_text": "（施工區域：完整一段，描述本次施工的確切區域範圍，至少 30 字）"
}}

注意：
1. 所有內容必須與報價單工程項目直接相關，要引用報價單中的具體材料名稱和規格
2. 問題說明必須描述「為什麼需要做這個工程」，不可寫成施工目的
3. 維修需求必須描述「這次要做什麼」，要能對應到報價項目
4. 施工內容是最關鍵的部分，每一條都要包含具體的材料、規格、數量、施作方式
5. 驗收標準必須可具體判定，不可使用「確保品質」等空泛語句
6. 驗收標準表至少要有 6 個類別
"""

# ============================================================
# QUOTE FILL PROMPTS (欣興 internal requisition form)
# ============================================================

QUOTE_FILL_SYSTEM_PROMPT = """\
你是一位資深的台灣工廠設備維護與工程管理專家。你的任務是根據報價單內容，填寫欣興電子內部的「保養、維修、改善、工程說明」表單。

此表單需要你以工廠內部人員的視角，用簡潔但完整的敘述來描述工程的現況、問題、對策和改善方式。

你的撰寫風格：
1. 簡潔扼要，使用條列式或短句
2. 用換行分隔不同要點（表格內使用換行）
3. 不可使用罐頭文字
4. 每個欄位都要與報價單內容直接相關
5. 金額要正確引用報價單數據

回傳格式必須是純 JSON。
"""

QUOTE_FILL_USER_PROMPT_TEMPLATE = """\
請根據以下報價單內容，生成欣興內部請購表單所需的各項內容。

## 工程資訊
- 工程名稱：{project_name}
- 文件編號：{document_number}
- 報價日期：{quote_date}
- 廠商名稱：{vendor_name}
- 施工位置：{project_location}
- 案子概略說明：{project_brief}
- 未稅金額：{subtotal}
- 含稅金額：{grand_total}

## 報價項目明細
{items_detail}

## 報價大項結構
{sections_detail}

## 要求輸出 JSON 格式

{{
  "purchase_name": "（請購名稱：工程名稱加報價單號，如 'YM1-5樓L區1K to 100級無塵室修改工程（LIQP2603015C）'）",
  "equipment_name": "（設備名稱：簡述工程類型，如 '無塵室修改工程'）",
  "station": "（站別：廠區代號，如 'YM1'）",
  "location": "（位置：具體樓層區域，如 '楊梅廠5樓L區'）",
  "cost_type": "（費用別：'修繕' 或 '非修繕'，新設或改善工程選非修繕）",
  "purchase_category_repair": "（請購類別(修繕)：若為非修繕則填 'NA'）",
  "purchase_category_non_repair": "（請購類別(非修繕)：如 '新設工程(20萬以上)'、'改善工程'等）",
  "work_reason": "（施作原因：如 '其他'、'設計缺失'、'品質異常'等）",
  "situation_desc_1": "（現況說明:1 - 描述主要工程的現況，用換行分多行，每行一個要點，4-5行。描述這是什麼案子、包含哪些主要工程內容）",
  "photo_location_1a": "（照片/施作位置 1a：主要施工區域名稱）",
  "photo_location_1b": "（照片/施作位置 1b：次要施工區域名稱）",
  "problem_risk_1": "（現況問題及風險 - 描述不施作會有什麼問題，用換行分多行，4-5行。描述功能不足、環境不符合需求等問題）",
  "countermeasure_1": "（對策評估說明 - 描述規劃的施工對策與順序，用換行分多行，4-5行）",
  "execution_1": "（執行改善/維修對策 - 描述具體執行方式，用換行分多行，4-5行。對應報價分項的執行方式）",
  "situation_desc_2": "（現況說明:2 - 描述其他費用項目的現況，如設計費、保險費、管理費等附加費用，用換行分多行）",
  "photo_location_2a": "（照片/施作位置 2a）",
  "photo_location_2b": "（照片/施作位置 2b）",
  "problem_risk_2": "（現況問題及風險 2 - 描述若未整體施作可能造成的介面衝突或系統不完整問題）",
  "improvement_desc": "（改善/維修說明 - 描述改善方式與整合做法）",
  "execution_2": "（執行改善/維修對策 2 - 描述驗收重點與確認項目）",
  "supplementary_notes": "（補充說明 - 包含報價單號、日期、工程名稱、金額明細，格式如：1. 報價單號：xxx；報價日期：xxx。\\n2. 工程名稱：xxx。\\n3. 未稅總計：xxx，稅額：xxx，含稅合計：xxx。）"
}}
"""

# ============================================================
# COMPLIANCE FILTER PROMPTS
# ============================================================

COMPLIANCE_SYSTEM_PROMPT = """\
你是一位台灣工程規範符合性判定專家。你的任務是根據報價單描述的工程內容，判斷一組既有的工程規範條目中，哪些與本次工程相關、哪些不相關。

判斷邏輯不能只做關鍵字比對，而要綜合考量：
- 工程類型（土建/機電/門禁/配管/清潔/設備維修/防護/水電/空調等）
- 材料性質
- 安裝方式
- 是否與本次項目實際相關

回傳格式必須是純 JSON。
"""

COMPLIANCE_USER_PROMPT_TEMPLATE = """\
## 本案工程概要
- 工程名稱：{project_name}
- 施工位置：{project_location}
- 案子概略說明：{project_brief}
- 工程項目摘要：
{items_summary}

## 既有規範條目清單
以下是待判定的規範條目，每一條都有一個 row_index（列索引）：
{spec_items_list}

## 要求輸出 JSON 格式

請針對每一條規範，判定是否與本案工程相關，回傳如下格式的 JSON 陣列：

[
  {{
    "row_index": 0,
    "is_relevant": true,
    "reason": "（簡述判定理由）"
  }},
  ...
]

注意：
1. 必須對每一條規範都給出判定結果
2. 相關性判定要基於工程實質內容，不能只看表面文字
3. 寧可多保留也不要誤刪（偏向保守判定）
"""

# ============================================================
# PROJECT BRIEF AUTO-GENERATION PROMPT
# ============================================================

PROJECT_BRIEF_SYSTEM_PROMPT = """\
你是一位台灣工程管理專家。根據報價單內容，生成一段案子概略說明。
回傳純文字，不要 JSON 格式。
"""

PROJECT_BRIEF_USER_PROMPT_TEMPLATE = """\
請根據以下報價單內容，生成一段 100-200 字的案子概略說明，簡述工程目的、主要施工範圍、涉及的系統與工程特點。

工程名稱：{project_name}
施工位置：{project_location}
報價項目：
{items_summary}

請直接回傳概略說明文字，不要加標題或編號。
"""


# ============================================================
# Helper functions
# ============================================================

def build_acceptance_prompt(quote_data, project_brief: str = "") -> tuple[str, str]:
    """Build the acceptance spec generation prompt from QuoteData."""
    items_detail = _format_items_detail(quote_data)
    terms_detail = _format_terms_detail(quote_data)

    brief = project_brief or quote_data.project_brief or quote_data.get_project_description()

    user_prompt = ACCEPTANCE_USER_PROMPT_TEMPLATE.format(
        project_name=quote_data.metadata.project_name,
        project_location=quote_data.metadata.project_location or "未指定",
        vendor_name=quote_data.metadata.vendor_name or "未指定",
        document_number=quote_data.metadata.document_number or "未指定",
        project_brief=brief,
        items_detail=items_detail,
        terms_detail=terms_detail,
    )

    return ACCEPTANCE_SYSTEM_PROMPT, user_prompt


def build_quote_fill_prompt(quote_data, project_brief: str = "") -> tuple[str, str]:
    """Build the quote fill (欣興 form) generation prompt."""
    items_detail = _format_items_detail(quote_data)
    brief = project_brief or quote_data.project_brief or quote_data.get_project_description()

    sections_detail = ""
    for section in quote_data.sections:
        sections_detail += f"- {section.section_number} {section.section_name}"
        if section.subtotal:
            sections_detail += f"（小計：{section.subtotal:,.0f}）"
        sections_detail += "\n"
        for item in section.items[:5]:
            sections_detail += f"  · {item.description}\n"
    if not sections_detail:
        sections_detail = "（未偵測到大項結構）"

    user_prompt = QUOTE_FILL_USER_PROMPT_TEMPLATE.format(
        project_name=quote_data.metadata.project_name,
        document_number=quote_data.metadata.document_number or "未指定",
        quote_date=str(quote_data.metadata.quote_date) if quote_data.metadata.quote_date else "未指定",
        vendor_name=quote_data.metadata.vendor_name or "未指定",
        project_location=quote_data.metadata.project_location or "未指定",
        project_brief=brief,
        subtotal=f"{quote_data.summary.subtotal:,.0f}",
        grand_total=f"{quote_data.summary.grand_total:,.0f}",
        items_detail=items_detail,
        sections_detail=sections_detail,
    )

    return QUOTE_FILL_SYSTEM_PROMPT, user_prompt


def build_compliance_prompt(
    quote_data, spec_items: list[dict], project_brief: str = ""
) -> tuple[str, str]:
    """Build the compliance filter prompt from QuoteData and spec items."""
    brief = project_brief or quote_data.project_brief or quote_data.get_project_description()

    items_summary = "\n".join(
        f"  - {item.description}"
        + (f"（{item.specification}）" if item.specification else "")
        for item in quote_data.get_real_items()
    )

    spec_items_list = "\n".join(
        f"  [{item['row_index']}] {item['content']}"
        for item in spec_items
    )

    user_prompt = COMPLIANCE_USER_PROMPT_TEMPLATE.format(
        project_name=quote_data.metadata.project_name,
        project_location=quote_data.metadata.project_location or "未指定",
        project_brief=brief,
        items_summary=items_summary or "（無法解析項目明細，請參考原始文字）",
        spec_items_list=spec_items_list,
    )

    return COMPLIANCE_SYSTEM_PROMPT, user_prompt


def build_project_brief_prompt(quote_data) -> tuple[str, str]:
    """Build the project brief auto-generation prompt."""
    items_summary = "\n".join(
        f"  - {item.description}"
        + (f"（{item.unit} x {item.quantity}）" if item.quantity > 1 else "")
        for item in quote_data.get_real_items()[:30]
    )

    user_prompt = PROJECT_BRIEF_USER_PROMPT_TEMPLATE.format(
        project_name=quote_data.metadata.project_name,
        project_location=quote_data.metadata.project_location or "未指定",
        items_summary=items_summary or "（無法解析項目明細）",
    )

    return PROJECT_BRIEF_SYSTEM_PROMPT, user_prompt


def _format_items_detail(quote_data) -> str:
    """Format line items into a readable string for the prompt."""
    lines = []
    for item in quote_data.get_real_items():
        parts = [f"項次 {item.seq}：{item.description}"]
        if item.specification:
            parts.append(f"  規格：{item.specification}")
        if item.brand:
            parts.append(f"  品牌：{item.brand}")
        if item.model:
            parts.append(f"  型號：{item.model}")
        parts.append(f"  單位：{item.unit}，數量：{item.quantity}")
        if item.unit_price:
            parts.append(f"  單價：{item.unit_price:,.0f}，複價：{item.total_price:,.0f}")
        if item.remark:
            parts.append(f"  備註：{item.remark}")
        if item.section:
            parts.append(f"  所屬大項：{item.section}")
        lines.append("\n".join(parts))

    return "\n\n".join(lines) if lines else "（無法解析報價項目明細，請根據以下原始文字推定）\n" + (quote_data.raw_text or "")[:3000]


def _format_terms_detail(quote_data) -> str:
    """Format terms and conditions for the prompt."""
    parts = []
    s = quote_data.summary
    if s.warranty_terms:
        parts.append(f"- 保固：{s.warranty_terms}")
    if s.insurance_terms:
        parts.append(f"- 保險：{s.insurance_terms}")
    if s.safety_terms:
        parts.append(f"- 安全：{s.safety_terms}")
    if s.cleanup_terms:
        parts.append(f"- 清運/清潔：{s.cleanup_terms}")
    if s.management_fee:
        parts.append(f"- 管理費：{s.management_fee}")
    if s.miscellaneous_fee:
        parts.append(f"- 運雜費：{s.miscellaneous_fee}")
    for note in s.notes:
        parts.append(f"- 備註：{note}")

    return "\n".join(parts) if parts else "（未偵測到附帶條件）"
