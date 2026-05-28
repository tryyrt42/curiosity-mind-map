"""
프로젝트: Curiosity Map (호기심 마인드맵)
설명: 단어를 던지면 꼬리에 꼬리를 무는 질문을 생성하고, Supabase에 캐싱하여 시각화하는 지식 탐험 툴
"""
import streamlit as st
import google.generativeai as genai
import json
from supabase import create_client, Client

# ==========================================
# 1. 페이지 세팅 및 전역 상태 초기화
# ==========================================
st.set_page_config(page_title="Curiosity Map", page_icon="🧠", layout="wide")

# 세션 상태 초기화
if "history" not in st.session_state:
    st.session_state["history"] = []  # 사용자가 지나온 단어들의 궤적
if "current_seed" not in st.session_state:
    st.session_state["current_seed"] = None
if "current_branches" not in st.session_state:
    st.session_state["current_branches"] = []
if "search_input" not in st.session_state:
    st.session_state["search_input"] = ""

# ==========================================
# 2. 환경 변수 및 DB 연결 (Supabase & Gemini)
# ==========================================
# 기존 Easy-Easy 프로젝트의 secrets.toml을 그대로 재활용합니다!
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

@st.cache_resource
def init_supabase() -> Client | None:
    if SUPABASE_URL and SUPABASE_KEY:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    return None

supabase = init_supabase()

if not GEMINI_API_KEY:
    st.error("🚨 `.streamlit/secrets.toml` 파일에 GEMINI_API_KEY를 설정해주세요.")
    st.stop()

genai.configure(api_key=GEMINI_API_KEY)

# ==========================================
# 3. 코어 로직: 캐싱 & 생성 엔진
# ==========================================
def fetch_curiosity_data(keyword: str):
    """
    1순위: Supabase 캐시 확인 (비용 0원, 속도 0.1초)
    2순위: 캐시에 없으면 Gemini API 호출 후 Supabase에 저장
    """
    keyword_clean = keyword.strip()
    
    # 1. Supabase 캐시 확인
    if supabase:
        try:
            # 💡 주의: Supabase에 'mindmap_cache' 테이블이 있어야 작동합니다.
            # 없으면 에러를 무시하고 2단계 API 호출로 넘어갑니다.
            response = supabase.table("mindmap_cache").select("branched_data").eq("keyword", keyword_clean).execute()
            if len(response.data) > 0:
                return response.data[0]["branched_data"]
        except Exception as e:
            pass # 테이블이 없거나 에러가 나면 API 호출로 넘어감

    # 2. Gemini API 호출 (Cache Miss)
    model = genai.GenerativeModel('gemini-1.5-flash', generation_config={"response_mime_type": "application/json"})
    
    prompt = f"""
    너는 인간의 호기심을 폭발시키는 지식 탐험 가이드야.
    사용자가 제시한 키워드: "{keyword_clean}"
    
    이 키워드에서 뻗어나갈 수 있는, 꼬리에 꼬리를 무는 본질적이고 흥미로운 질문 3가지를 만들어.
    "왜 그럴까?", "어떤 원리일까?", "무엇이 다를까?" 등 깊게 파고드는 질문이어야 해.
    
    반드시 아래의 JSON 구조로만 답변해:
    {{
        "seed": "{keyword_clean}",
        "branches": [
            {{"question": "파생 질문 1", "hint": "질문에 대한 핵심을 찌르는 1~2줄의 명쾌한 힌트"}},
            {{"question": "파생 질문 2", "hint": "질문에 대한 핵심을 찌르는 1~2줄의 명쾌한 힌트"}},
            {{"question": "파생 질문 3", "hint": "질문에 대한 핵심을 찌르는 1~2줄의 명쾌한 힌트"}}
        ]
    }}
    """
    
    try:
        api_response = model.generate_content(prompt)
        data = json.loads(api_response.text)
        
        # 3. 생성된 데이터를 Supabase에 캐싱 (다음 유저를 위해)
        if supabase and data:
            try:
                supabase.table("mindmap_cache").insert({
                    "keyword": keyword_clean, 
                    "branched_data": data
                }).execute()
            except Exception as e:
                pass # 테이블이 아직 없으면 무시
                
        return data
    except Exception as e:
        st.error(f"AI 생성 오류: {e}")
        return None

# ==========================================
# 4. 전역 디자인 (CSS)
# ==========================================
st.markdown("""
<style>
    .stApp { background-color: #0f172a; color: #f8fafc; }
    h1 { background: linear-gradient(90deg, #facc15, #f43f5e); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 900 !important; text-align: center; font-size: 3.5rem; margin-bottom: 0;}
    .subtitle { text-align: center; color: #94a3b8; font-size: 1.2rem; margin-bottom: 3rem; }
    
    /* 카드 디자인 */
    div[data-testid="stVerticalBlock"] > div > div { border-radius: 12px; }
    .question-card {
        background: rgba(30, 41, 59, 0.7);
        border: 1px solid rgba(244, 63, 94, 0.3);
        border-radius: 12px;
        padding: 24px;
        margin-top: 10px;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .question-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 20px rgba(244, 63, 94, 0.2);
    }
    .card-hint { color: #cbd5e1; font-size: 0.95rem; margin-bottom: 20px; line-height: 1.6; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 5. UI: 헤더 및 검색창
# ==========================================
st.markdown("<h1>🧠 Curiosity Map</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>하나의 키워드에서 뻗어나가는 지식의 무한 루프</div>", unsafe_allow_html=True)

col_spacer1, col_search, col_spacer2 = st.columns([1, 2, 1])

def submit_search():
    st.session_state["current_seed"] = st.session_state["search_input"]

with col_search:
    search_query = st.text_input(
        "탐험할 키워드나 궁금한 점을 던져보세요!", 
        placeholder="예: 블랙홀, 자본주의, 양자역학, 얼음", 
        key="search_input",
        on_change=submit_search
    )

# 버튼 클릭이나 엔터로 시드(Seed)가 설정되었을 때
target_seed = st.session_state["current_seed"]

if target_seed:
    with st.spinner(f"✨ '{target_seed}'에 대한 호기심을 연결하고 있습니다..."):
        result_data = fetch_curiosity_data(target_seed)
        
        if result_data:
            st.session_state["current_branches"] = result_data.get("branches", [])
            # 히스토리에 없으면 추가 (뒤로가기 등 구현용)
            if not st.session_state["history"] or st.session_state["history"][-1] != target_seed:
                st.session_state["history"].append(target_seed)

# ==========================================
# 6. UI: 결과 시각화 (마인드맵 + 인터랙티브 카드)
# ==========================================
if st.session_state["current_seed"] and st.session_state["current_branches"]:
    st.divider()
    
    seed = st.session_state["current_seed"]
    branches = st.session_state["current_branches"]
    
    # --- [A] 시각적 마인드맵 (Mermaid.js) ---
    st.markdown("<h3 style='text-align: center; margin-bottom: 20px;'>🗺️ 생각의 지도</h3>", unsafe_allow_html=True)
    
    mermaid_code = f"graph LR\n"
    mermaid_code += f"    A(({seed}))\n"
    mermaid_code += f"    style A fill:#f43f5e,stroke:#fff,stroke-width:2px,color:#fff\n"
    
    for i, branch in enumerate(branches):
        q_text = branch['question'].replace('?', '?<br>')
        mermaid_code += f"    A --> B{i}([{q_text}])\n"
        mermaid_code += f"    style B{i} fill:#1e293b,stroke:#facc15,stroke-width:2px,color:#fff\n"
    
    import streamlit.components.v1 as components
    mermaid_html = f"""
    <div class="mermaid" style="display: flex; justify-content: center;">
        {mermaid_code}
    </div>
    <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
    <script>mermaid.initialize({{startOnLoad:true, theme: 'dark'}});</script>
    """
    components.html(mermaid_html, height=300)
    
    # --- [B] 파생 질문 카드 ---
    st.markdown("### 🔍 더 깊이 파고들기")
    
    cols = st.columns(3)
    for i, branch in enumerate(branches):
        with cols[i]:
            st.markdown(f"""
            <div class="question-card">
                <div>
                    <h4 style='color: #facc15; margin-top: 0;'>{branch['question']}</h4>
                    <div class='card-hint'>💡 {branch['hint']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # 카드 하단에 꼬리물기 버튼 배치
            # 콜백을 사용하여 버튼 클릭 시 해당 질문을 새로운 시드로 설정
            def dive_deeper(new_seed=branch['question']):
                st.session_state["search_input"] = new_seed
                st.session_state["current_seed"] = new_seed

            st.button(
                "➡️ 이 질문으로 계속 파기", 
                key=f"dive_{i}_{seed}", 
                use_container_width=True,
                on_click=dive_deeper
            )
