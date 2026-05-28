"""
프로젝트: Curiosity Map (호기심 캔버스 마인드맵) - V2 완벽 수정본
특징: 
1. 캔버스 정중앙에 위치한 단일 텍스트 입력창 (자동완성 드롭다운 제거)
2. 페이지 전환 없이, 입력창이 있던 그 자리에 즉시 마인드맵 노드가 생성됨
3. 물리 엔진(Physics)을 활용해 우측으로 스무스하게 뻗어나가는 곡선 애니메이션
"""
import streamlit as st
import google.generativeai as genai
import json
from supabase import create_client, Client
from streamlit_agraph import agraph, Node, Edge, Config

# ==========================================
# 1. 페이지 세팅 및 전역 상태 초기화
# ==========================================
st.set_page_config(page_title="Curiosity Canvas", page_icon="🌌", layout="wide")

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
if "pending_keyword" not in st.session_state:
    st.session_state.pending_keyword = None

# ==========================================
# 2. 환경 변수 및 DB 연결
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
    너는 인간의 호기심을 폭발시키는 지식 탐험 가이드야.
    사용자가 클릭하거나 입력한 키워드: "{keyword_clean}"
    
    이 키워드에서 뻗어나갈 수 있는 꼬리에 꼬리를 무는 본질적인 질문 3가지를 만들어.
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
                supabase.table("mindmap_cache").insert({
                    "keyword": keyword_clean, 
                    "branched_data": data
                }).execute()
            except Exception:
                pass 
                
        return data
    except Exception as e:
        st.error(f"AI 생성 오류: {e}")
        return None

def expand_graph(keyword: str, parent_id: str):
    """주어진 키워드로 3개의 질문 상자를 만들고 곡선으로 연결합니다."""
    data = fetch_curiosity_data(keyword)
    if data and "branches" in data:
        for idx, branch in enumerate(data["branches"]):
            q = branch["question"]
            node_id = f"{parent_id}_||_{idx}_||_{q}" 
            
            # 3번 사진 감성에 맞춘 심플한 화이트 박스 디자인
            st.session_state.nodes.append(Node(
                id=node_id,
                label=q,
                shape="box",
                color={"border": "#cbd5e1", "background": "#ffffff", "highlight": {"border": "#6366f1", "background": "#f8fafc"}},
                font={"color": "#333", "size": 15},
                margin={"top": 15, "bottom": 15, "left": 25, "right": 25}
            ))
            
            # 유려하게 뻗어나가는 곡선 (Cubic Bezier)
            st.session_state.edges.append(Edge(
                source=parent_id,
                target=node_id,
                color="#cbd5e1",
                width=2,
                smooth={"type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.6}
            ))
            
        st.session_state.expanded_nodes.add(parent_id)

# ==========================================
# 4. 동적 UI (In-Place 렌더링)
# ==========================================

# 캔버스 전체 여백 제거 및 완벽한 중앙 정렬을 위한 전역 CSS
st.markdown("""
<style>
    .stApp { background-color: #fafafa; }
    header { visibility: hidden; } /* 상단 여백 완전 제거 */
    
    /* 메인 블록의 불필요한 패딩 제거 */
    .main .block-container {
        padding: 0 !important;
        max-width: 100% !important;
        margin: 0 !important;
    }
    
    /* 🚨 입력창 완벽한 정중앙 고정 핵 (드롭다운 방지) */
    div[data-testid="stTextInput"] {
        position: fixed;
        top: 45%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 350px !important;
        z-index: 9999;
    }
    div[data-baseweb="input"] > div {
        background-color: #ffffff !important;
        border: 2px solid #cbd5e1 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05) !important;
        transition: all 0.3s ease;
    }
    div[data-baseweb="input"] > div:focus-within {
        border-color: #6366f1 !important;
        box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.2) !important;
    }
    input {
        text-align: center !important;
        font-size: 1.2rem !important;
        padding: 1.2rem !important;
        color: #1e293b !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# [Phase 1] 빈 캔버스 + 중앙 입력창
# ---------------------------------------------------------
if not st.session_state.is_started:
    # autocomplete="off"로 기존 입력 기록 드롭다운 끔
    seed_word = st.text_input("", placeholder="궁금한 단어를 입력하세요", key="init_seed", label_visibility="collapsed", autocomplete="off")
    
    if seed_word:
        # 엔터를 치는 순간 텍스트 박스는 사라지고, 그 자리에 그래프 루트 노드 생성
        st.session_state.is_started = True
        st.session_state.nodes.append(Node(
            id=seed_word,
            label=seed_word,
            shape="box",
            color={"border": "#6366f1", "background": "#ffffff"},
            font={"color": "#1e293b", "size": 18, "bold": True},
            margin={"top": 15, "bottom": 15, "left": 30, "right": 30}
        ))
        # 백그라운드 렌더링 지시
        st.session_state.pending_expansion = seed_word
        st.session_state.pending_keyword = seed_word
        st.rerun()

# ---------------------------------------------------------
# [Phase 2] 인플레이스 마인드맵 (화면 전환 없음)
# ---------------------------------------------------------
else:
    # 트리 구조 (좌 -> 우) 물리 엔진 설정
    config = Config(
        width="100%",
        height=900, # 꽉 찬 높이
        directed=True, 
        physics=True, # 🌟 노드가 스프링처럼 튀어나오는 물리 효과
        hierarchical=True,
        layout={
            "hierarchical": {
                "enabled": True,
                "direction": "LR",        # 왼쪽에서 오른쪽으로 뻗어나감
                "sortMethod": "directed", 
                "levelSeparation": 350,   # 좌우 화살표 길이 (스무스한 선)
                "nodeSpacing": 120        # 위아래 상자 간격
            }
        },
        interaction={"hover": True, "dragNodes": False}
    )
    
    # 캔버스에 마인드맵 그리기
    clicked_node_id = agraph(nodes=st.session_state.nodes, edges=st.session_state.edges, config=config)
    
    # [백그라운드 통신] 상자가 새로 생기거나 클릭되었을 때 스무스하게 뻗어 나오는 로직
    if st.session_state.pending_expansion:
        node_id = st.session_state.pending_expansion
        keyword = st.session_state.pending_keyword
        
        with st.spinner("생각의 가지를 뻗는 중..."):
            expand_graph(keyword, node_id)
            
        st.session_state.pending_expansion = None
        st.session_state.pending_keyword = None
        st.rerun() # 새로운 상자들이 우측으로 밀려 나오는 애니메이션 트리거
        
    # 이미 띄워진 상자를 유저가 클릭했을 때
    elif clicked_node_id and clicked_node_id not in st.session_state.expanded_nodes:
        keyword = clicked_node_id.split("_||_")[-1] if "_||_" in clicked_node_id else clicked_node_id
        st.session_state.pending_expansion = clicked_node_id
        st.session_state.pending_keyword = keyword
        st.rerun()

    # 초기화 버튼
    st.markdown("""
        <style>
        .reset-btn-container { position: fixed; bottom: 30px; left: 30px; z-index: 9999; }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("<div class='reset-btn-container'>", unsafe_allow_html=True)
    if st.button("🔄 캔버스 비우기"):
        st.session_state.nodes = []
        st.session_state.edges = []
        st.session_state.expanded_nodes = set()
        st.session_state.is_started = False
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
