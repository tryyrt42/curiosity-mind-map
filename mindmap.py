"""
프로젝트: Curiosity Map (호기심 캔버스 마인드맵)
특징: 광활한 빈 공간의 중앙 입력창, 스무스하게 우측(LR)으로 뻗어나가는 인터랙티브 노드 UI
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
# 3. 호기심 생성 엔진 (캐싱 포함)
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
    제시된 키워드/질문: "{keyword_clean}"
    
    이 주제에서 뻗어나갈 수 있는, 꼬리에 꼬리를 무는 본질적이고 흥미로운 질문 3가지를 만들어.
    반드시 아래의 JSON 구조로만 답변해:
    {{
        "branches": [
            {{"question": "파생 질문 1", "hint": "마우스를 올렸을 때 보여줄 1~2줄의 명쾌한 해답/힌트"}},
            {{"question": "파생 질문 2", "hint": "마우스를 올렸을 때 보여줄 1~2줄의 명쾌한 해답/힌트"}},
            {{"question": "파생 질문 3", "hint": "마우스를 올렸을 때 보여줄 1~2줄의 명쾌한 해답/힌트"}}
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
    """주어진 키워드로 질문을 생성하고 노드/엣지를 추가합니다."""
    data = fetch_curiosity_data(keyword)
    if data and "branches" in data:
        for idx, branch in enumerate(data["branches"]):
            q = branch["question"]
            h = branch["hint"]
            node_id = f"{parent_id}_||_{idx}_||_{q}" 
            
            # 파생 질문 노드 추가 (클릭 가능한 예쁜 박스)
            st.session_state.nodes.append(Node(
                id=node_id,
                label=q,
                title=f"💡 {h}", # 마우스 오버 시 힌트 툴팁 표시
                shape="box",
                color={"border": "#f43f5e", "background": "#1e293b", "highlight": {"border": "#facc15", "background": "#334155"}},
                font={"color": "white", "size": 16},
                margin={"top": 15, "bottom": 15, "left": 20, "right": 20}
            ))
            
            # 스무스한 곡선 엣지 추가
            st.session_state.edges.append(Edge(
                source=parent_id,
                target=node_id,
                color="#8b5cf6",
                width=3,
                smooth={"type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.4}
            ))
            
        st.session_state.expanded_nodes.add(parent_id)

# ==========================================
# 4. 동적 UI (빈 캔버스 vs 마인드맵)
# ==========================================

# 기본 우주 배경 CSS
st.markdown("""
<style>
    .stApp {
        background-color: #050505;
        background-image: radial-gradient(circle at center, #111 0%, #000 100%);
    }
    header { visibility: hidden; } /* 상단 여백 제거 */
</style>
""", unsafe_allow_html=True)

if not st.session_state.is_started:
    # ---------------------------------------------------------
    # [1단계] 광활한 공간, 정중앙에 빛나는 검색창만 있는 상태
    # ---------------------------------------------------------
    st.markdown("""
    <style>
        /* 화면 정중앙 배치 해킹 */
        div[data-testid="stVerticalBlock"] {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 85vh;
        }
        
        /* 💡 빛나는 인풋 박스 디자인 */
        div[data-baseweb="input"] > div {
            background-color: #0b0f19 !important;
            border: 2px solid #8b5cf6 !important;
            box-shadow: 0 0 40px rgba(139, 92, 246, 0.6), inset 0 0 20px rgba(139, 92, 246, 0.2) !important;
            border-radius: 25px !important;
            transition: all 0.3s ease;
            width: 600px;
        }
        div[data-baseweb="input"] > div:hover {
            box-shadow: 0 0 60px rgba(139, 92, 246, 0.8), inset 0 0 30px rgba(139, 92, 246, 0.4) !important;
        }
        input {
            color: #fff !important;
            font-size: 2rem !important;
            text-align: center !important;
            padding: 1.5rem !important;
        }
    </style>
    """, unsafe_allow_html=True)

    seed_word = st.text_input("", placeholder="얼음, 양자역학, 인공지능...", key="init_seed", label_visibility="collapsed")
    
    if seed_word:
        with st.spinner("호기심의 우주를 여는 중..."):
            # 씨앗 노드 생성
            st.session_state.nodes.append(Node(
                id=seed_word,
                label=seed_word,
                shape="box",
                color={"border": "#8b5cf6", "background": "#0b0f19"},
                font={"color": "white", "size": 24, "bold": True},
                margin={"top": 20, "bottom": 20, "left": 30, "right": 30}
            ))
            # 파생 질문 가지치기
            expand_graph(seed_word, seed_word)
            st.session_state.is_started = True
            st.rerun()

else:
    # ---------------------------------------------------------
    # [2단계] 마인드맵 캔버스 뷰 (클릭하며 무한 확장)
    # ---------------------------------------------------------
    st.markdown("""
    <style>
        /* 캔버스 뷰에서는 위쪽 빈 공간 제거하고 화면 전체 사용 */
        div[data-testid="stVerticalBlock"] {
            display: block;
            height: auto;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # 좌에서 우(LR)로 뻗어나가는 계층형(Hierarchical) 레이아웃 설정
    config = Config(
        width="100%",
        height=800,
        directed=True, 
        physics=False, 
        hierarchical=True,
        direction="LR",      # 🔥 무조건 왼쪽에서 오른쪽으로 뻗어나감
        nodeSpacing=120,     # 노드 간 세로 간격
        levelSeparation=400, # 노드 간 가로 길이 (스무스한 선의 길이)
        interaction={"hover": True} # 마우스 올리면 힌트 툴팁 보이게
    )
    
    # 화면에 그래프 렌더링 (클릭 시 노드 ID 반환)
    clicked_node_id = agraph(nodes=st.session_state.nodes, edges=st.session_state.edges, config=config)
    
    # 노드를 클릭했고, 아직 가지치기를 안 한 노드라면? -> 또 가지치기!
    if clicked_node_id and clicked_node_id not in st.session_state.expanded_nodes:
        # 노드 ID에서 실제 질문 텍스트만 추출 (구분자 '_||_' 사용)
        actual_question = clicked_node_id.split("_||_")[-1] if "_||_" in clicked_node_id else clicked_node_id
        
        with st.spinner(f"'{actual_question}'에 대해 더 깊이 파고드는 중..."):
            expand_graph(actual_question, clicked_node_id)
            st.rerun()
            
    # 초기화 버튼
    if st.button("🚀 새 우주 열기 (초기화)"):
        st.session_state.nodes = []
        st.session_state.edges = []
        st.session_state.expanded_nodes = set()
        st.session_state.is_started = False
        st.rerun()
