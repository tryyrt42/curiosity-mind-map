"""
프로젝트: Curiosity Map (호기심 캔버스 마인드맵) - 광활한 무한 캔버스 버전
"""
import streamlit as st
import google.generativeai as genai
import json
import uuid
from supabase import create_client, Client
from streamlit_agraph import agraph, Node, Edge, Config

# ==========================================
# 1. 페이지 세팅 (여백 없는 완벽한 전체화면)
# ==========================================
st.set_page_config(page_title="Curiosity Canvas", page_icon="🌌", layout="wide")

st.markdown("""
<style>
    /* 전체 배경을 은은한 도트 감성의 빈 공간으로 */
    .stApp { background-color: #f8fafc; }
    
    /* 화면을 좁고 길쭉하게 만드는 주범(여백, 헤더, 푸터) 완벽 제거 */
    header { display: none !important; }
    footer { display: none !important; }
    .main .block-container {
        padding: 0 !important;
        max-width: 100% !important;
        margin: 0 !important;
    }
    
    /* 1단계: 정중앙 글상자 디자인 */
    .center-container {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
    }
    div[data-testid="stTextInput"] {
        width: 250px !important;
    }
    div[data-baseweb="input"] > div {
        background-color: #ffffff !important;
        border: 2px solid #cbd5e1 !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05) !important;
    }
    div[data-baseweb="input"] > div:focus-within {
        border-color: #8b5cf6 !important;
        box-shadow: 0 0 0 4px rgba(139, 92, 246, 0.1) !important;
    }
    input {
        text-align: center !important;
        font-size: 1.2rem !important;
        padding: 1.2rem !important;
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
if "is_started" not in st.session_state:
    st.session_state.is_started = False
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
    키워드: "{keyword_clean}"
    이 키워드에서 파생되는 심도 있는 질문 3가지를 만들어.
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
    except Exception:
        return None

def expand_graph(keyword: str, parent_id: str):
    data = fetch_curiosity_data(keyword)
    if data and "branches" in data:
        for idx, branch in enumerate(data["branches"]):
            q = branch["question"]
            node_id = f"{parent_id}_||_{idx}_||_{q}" 
            
            # 사각형 상자 디자인 (widthConstraint로 텍스트 자동 줄바꿈)
            st.session_state.nodes.append(Node(
                id=node_id,
                label=q,
                shape="box",
                color={"border": "#cbd5e1", "background": "#ffffff", "highlight": {"border": "#8b5cf6", "background": "#f8fafc"}},
                font={"color": "#1e293b", "size": 15},
                widthConstraint={"maximum": 250}, 
                margin={"top": 12, "bottom": 12, "left": 15, "right": 15}
            ))
            
            # 스무스한 곡선 연결선
            st.session_state.edges.append(Edge(
                source=parent_id,
                target=node_id,
                color="#cbd5e1",
                width=2,
                smooth={"type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.6}
            ))
            
        st.session_state.expanded_nodes.add(parent_id)

# ==========================================
# 4. 동적 UI 렌더링
# ==========================================

if not st.session_state.is_started:
    # ---------------------------------------------------------
    # [1단계] 광활한 캔버스 정중앙의 글상자
    # ---------------------------------------------------------
    st.markdown("<div class='center-container'>", unsafe_allow_html=True)
    
    # 💡 난수(uuid)를 key에 섞어 브라우저의 과거 검색어 자동완성을 강제로 무력화시킵니다.
    seed_word = st.text_input(
        "시드", 
        placeholder="단어를 입력하세요", 
        label_visibility="collapsed",
        key=f"seed_{uuid.uuid4()}" 
    )
    st.markdown("</div>", unsafe_allow_html=True)
    
    if seed_word:
        with st.spinner("생각을 그리는 중..."):
            st.session_state.nodes.append(Node(
                id=seed_word,
                label=seed_word,
                shape="box",
                color={"border": "#8b5cf6", "background": "#f5f3ff"},
                font={"color": "#1e293b", "size": 18, "bold": True},
                widthConstraint={"maximum": 200},
                margin={"top": 15, "bottom": 15, "left": 20, "right": 20}
            ))
            st.session_state.is_started = True
            st.session_state.pending_expansion = seed_word
            st.rerun()

else:
    # ---------------------------------------------------------
    # [2단계] 스르륵 그려지는 무한 마인드맵 캔버스
    # ---------------------------------------------------------
    config = Config(
        width="100%",
        height="100vh", # 💡 찌그러짐 방지: 화면 전체 높이 사용
        directed=True, 
        physics=True,   # 💡 노드가 스르륵 날아와서 자리잡는 물리 애니메이션 효과
        hierarchical=True,
        layout={
            "hierarchical": {
                "enabled": True,
                "direction": "LR",        # 완벽한 왼쪽 -> 오른쪽 방향
                "sortMethod": "directed",
                "levelSeparation": 300,   # 화살표 길이
                "nodeSpacing": 100        # 상자 위아래 간격
            }
        },
        interaction={"hover": True, "dragNodes": True}
    )
    
    # 캔버스 렌더링 (클릭한 상자 ID 반환)
    clicked_node_id = agraph(nodes=st.session_state.nodes, edges=st.session_state.edges, config=config)
    
    # 백그라운드 확장 처리 (새로운 가지 치기)
    if st.session_state.pending_expansion:
        node_id = st.session_state.pending_expansion
        keyword = node_id.split("_||_")[-1] if "_||_" in node_id else node_id
        
        with st.spinner("생각을 확장하는 중..."):
            expand_graph(keyword, node_id)
            
        st.session_state.pending_expansion = None
        st.rerun() # 새로운 상자들이 우측으로 '스르륵' 등장함
        
    # 이미 띄워진 상자를 유저가 클릭했을 때
    elif clicked_node_id and clicked_node_id not in st.session_state.expanded_nodes:
        st.session_state.pending_expansion = clicked_node_id
        st.rerun()

    # 초기화 버튼 (좌측 하단 구석에 조그맣게)
    st.markdown("""
        <style>
        .reset-btn { position: fixed; bottom: 20px; left: 20px; z-index: 999; opacity: 0.5; transition: 0.3s; }
        .reset-btn:hover { opacity: 1; }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("<div class='reset-btn'>", unsafe_allow_html=True)
    if st.button("🔄 리셋"):
        st.session_state.nodes = []
        st.session_state.edges = []
        st.session_state.expanded_nodes = set()
        st.session_state.is_started = False
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
