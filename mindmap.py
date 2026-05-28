"""
프로젝트: Curiosity Map (호기심 캔버스 마인드맵) - V3 (Miro/Obsidian Style)
특징: 
1. 화면 전환 없음: 처음부터 100% 꽉 찬 무한 캔버스 로드
2. 텍스트 자동 줄바꿈: 노드가 길쭉해지지 않고 예쁜 직사각형 유지 (widthConstraint)
3. 스무스한 물리 엔진: 노드 생성 시 겹치지 않고 밀어내며 자연스럽게 배치됨
"""
import streamlit as st
import google.generativeai as genai
import json
from supabase import create_client, Client
from streamlit_agraph import agraph, Node, Edge, Config

# ==========================================
# 1. 페이지 세팅 (여백 없는 완벽한 전체화면)
# ==========================================
st.set_page_config(page_title="Curiosity Canvas", page_icon="🌌", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #fafafa; }
    /* 상단 헤더, 여백 완전히 날리기 */
    header { visibility: hidden; }
    .main .block-container {
        padding: 0rem !important;
        max-width: 100% !important;
        margin: 0 !important;
    }
    /* 하단 입력창을 커맨드 팔레트처럼 플로팅 */
    .floating-input {
        position: fixed;
        bottom: 30px;
        left: 50%;
        transform: translateX(-50%);
        width: 600px;
        z-index: 99999;
        background: white;
        padding: 15px 25px;
        border-radius: 50px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
    }
</style>
""", unsafe_allow_html=True)

# 상태 초기화
if "nodes" not in st.session_state:
    st.session_state.nodes = []
if "edges" not in st.session_state:
    st.session_state.edges = []
if "expanded_nodes" not in st.session_state:
    st.session_state.expanded_nodes = set()
if "pending_expansion" not in st.session_state:
    st.session_state.pending_expansion = None 

# ==========================================
# 2. API 세팅
# ==========================================
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
    st.error("🚨 Secrets에 GEMINI_API_KEY를 설정해주세요.")
    st.stop()

genai.configure(api_key=GEMINI_API_KEY)

# ==========================================
# 3. 호기심 생성 엔진
# ==========================================
def fetch_curiosity_data(keyword: str):
    keyword_clean = keyword.strip()
    
    if supabase:
        try:
            response = supabase.table("mindmap_cache").select("branched_data").eq("keyword", keyword_clean).execute()
            if len(response.data) > 0:
                return response.data[0]["branched_data"]
        except Exception:
            pass 

    model = genai.GenerativeModel('gemini-3.1-flash-lite', generation_config={"response_mime_type": "application/json"})
    
    prompt = f"""
    너는 인간의 호기심을 폭발시키는 지식 툴이야.
    키워드: "{keyword_clean}"
    이 키워드에서 파생되는 심도 있는 질문 3가지를 만들어. 너무 길지 않게 핵심만 작성해.
    반드시 아래의 JSON 구조로만 답변해:
    {{
        "branches": [
            {{"question": "파생 질문 1"}},
            {{"question": "파생 질문 2"}},
            {{"question": "파생 질문 3"}}
        ]
    }}
    """
    try:
        api_response = model.generate_content(prompt)
        data = json.loads(api_response.text)
        if supabase and data:
            try:
                supabase.table("mindmap_cache").insert({"keyword": keyword_clean, "branched_data": data}).execute()
            except Exception:
                pass 
        return data
    except Exception as e:
        return None

def expand_graph(keyword: str, parent_id: str):
    data = fetch_curiosity_data(keyword)
    if data and "branches" in data:
        for idx, branch in enumerate(data["branches"]):
            q = branch["question"]
            node_id = f"{parent_id}_||_{idx}_||_{q}" 
            
            st.session_state.nodes.append(Node(
                id=node_id,
                label=q,
                shape="box",
                color={"border": "#cbd5e1", "background": "#ffffff", "highlight": {"border": "#8b5cf6", "background": "#f8fafc"}},
                font={"color": "#1e293b", "size": 15},
                # 💡 핵심: 상자가 길쭉해지지 않도록 최대 너비 제한 (자동 줄바꿈 됨)
                widthConstraint={"maximum": 220}, 
                margin={"top": 12, "bottom": 12, "left": 15, "right": 15}
            ))
            
            st.session_state.edges.append(Edge(
                source=parent_id,
                target=node_id,
                color="#cbd5e1",
                width=2,
                smooth={"type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.5}
            ))
            
        st.session_state.expanded_nodes.add(parent_id)

# ==========================================
# 4. 캔버스 레이아웃 및 렌더링
# ==========================================

# 💡 마이로/옵시디언 감성의 정밀한 물리엔진 세팅
config = Config(
    width="100%",
    height="100vh", # 화면 꽉 채우기
    directed=True, 
    physics={
        "hierarchicalRepulsion": {
            "nodeDistance": 180, # 노드 간격 (겹침 방지)
            "springLength": 250, # 화살표 길이
            "springConstant": 0.05
        },
        "solver": "hierarchicalRepulsion"
    },
    layout={
        "hierarchical": {
            "enabled": True,
            "direction": "LR",      # 좌에서 우로
            "sortMethod": "directed",
            "levelSeparation": 300, # 가로 단계별 간격
        }
    },
    interaction={"hover": True, "dragNodes": True, "zoomView": True}
)

# 캔버스 렌더링 (화면에는 항상 이 캔버스가 떠 있습니다)
clicked_node_id = agraph(nodes=st.session_state.nodes, edges=st.session_state.edges, config=config)

# 상단 혹은 하단에 고정된 검색창 
st.markdown("<div class='floating-input'>", unsafe_allow_html=True)
with st.form("seed_form", clear_on_submit=True, border=False):
    col1, col2 = st.columns([5, 1])
    with col1:
        seed_word = st.text_input("🌱 시작 키워드 입력", placeholder="예: 양자역학, 자본주의...", label_visibility="collapsed")
    with col2:
        submit_btn = st.form_submit_button("확장")
st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 5. 인터랙션 처리 (키워드 입력 OR 노드 클릭)
# ==========================================
if submit_btn and seed_word:
    # 캔버스 비우고 정중앙에 루트 노드 생성
    st.session_state.nodes = [Node(
        id=seed_word,
        label=seed_word,
        shape="box",
        color={"border": "#8b5cf6", "background": "#f5f3ff"},
        font={"color": "#1e293b", "size": 18, "bold": True},
        widthConstraint={"maximum": 200},
        margin={"top": 15, "bottom": 15, "left": 20, "right": 20}
    )]
    st.session_state.edges = []
    st.session_state.expanded_nodes = set()
    st.session_state.pending_expansion = seed_word
    st.rerun()

# 백그라운드 확장 처리 (루트 노드 생성 직후)
if st.session_state.pending_expansion:
    node_id = st.session_state.pending_expansion
    keyword = node_id.split("_||_")[-1] if "_||_" in node_id else node_id
    
    expand_graph(keyword, node_id)
    st.session_state.pending_expansion = None
    st.rerun()

# 기존에 띄워진 노드를 클릭했을 때 확장
if clicked_node_id and clicked_node_id not in st.session_state.expanded_nodes:
    st.session_state.pending_expansion = clicked_node_id
    st.rerun()
