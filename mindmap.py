"""
프로젝트: Curiosity Map (호기심 캔버스 마인드맵)
특징: 
1. 초기 화면: 텅 빈 공간 정중앙에 위치한 작은 입력 상자
2. 레이아웃: 완벽한 Left-to-Right 계층 구조 (3번 사진 완벽 구현)
3. 인터랙션: 노드 클릭 시 해당 노드를 뿌리로 삼아 3개의 새 가지가 무한 확장
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

    # 💡 최신 모델 적용
    model = genai.GenerativeModel('gemini-3.1-flash-lite', generation_config={"response_mime_type": "application/json"})
    
    prompt = f"""
    너는 인간의 호기심을 폭발시키는 지식 탐험 가이드야.
    사용자가 클릭하거나 입력한 키워드: "{keyword_clean}"
    
    이 키워드에서 뻗어나갈 수 있는 파생 질문 3가지를 만들어.
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
    """주어진 키워드로 질문을 생성하고 노드/엣지를 우측으로 추가합니다."""
    data = fetch_curiosity_data(keyword)
    if data and "branches" in data:
        for idx, branch in enumerate(data["branches"]):
            q = branch["question"]
            node_id = f"{parent_id}_||_{idx}_||_{q}" 
            
            # 파생 질문 노드 추가
            st.session_state.nodes.append(Node(
                id=node_id,
                label=q,
                shape="box",
                # 용규님 사진 감성에 맞춘 심플하고 세련된 화이트/그레이톤
                color={"border": "#ccc", "background": "#ffffff", "highlight": {"border": "#8b5cf6", "background": "#f5f3ff"}},
                font={"color": "#333", "size": 14},
                margin={"top": 15, "bottom": 15, "left": 20, "right": 20}
            ))
            
            # 3번 사진과 동일한 스무스한 곡선 엣지
            st.session_state.edges.append(Edge(
                source=parent_id,
                target=node_id,
                color="#b4b4b4",
                width=2,
                # 완벽한 베지어 곡선 설정
                smooth={"type": "cubicBezier", "forceDirection": "horizontal", "roundness": 0.5}
            ))
            
        st.session_state.expanded_nodes.add(parent_id)

# ==========================================
# 4. 동적 UI (빈 캔버스 vs 마인드맵)
# ==========================================

# 기본 백그라운드 색상 (깔끔한 도트 배경 느낌을 위해 연한 회색)
st.markdown("""
<style>
    .stApp { background-color: #fafafa; }
    header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

if not st.session_state.is_started:
    # ---------------------------------------------------------
    # [1단계] 중앙에 작고 예쁜 입력창 하나만 덩그러니
    # ---------------------------------------------------------
    st.markdown("""
    <style>
        /* 컨테이너를 화면 정중앙에 고정 */
        .center-box {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 70vh;
        }
        
        /* 길쭉한 기본 input을 작고 예쁜 상자로 변경 */
        div[data-baseweb="input"] {
            width: 200px !important; /* 가로폭 축소 */
            margin: 0 auto;
        }
        div[data-baseweb="input"] > div {
            background-color: #ffffff !important;
            border: 2px solid #e2e8f0 !important;
            border-radius: 8px !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1) !important;
            transition: all 0.3s ease;
        }
        div[data-baseweb="input"] > div:focus-within {
            border-color: #8b5cf6 !important;
            box-shadow: 0 0 15px rgba(139, 92, 246, 0.4) !important; /* 은은한 빛 효과 */
        }
        input {
            color: #333 !important;
            font-size: 1.1rem !important;
            text-align: center !important;
            padding: 1rem !important;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<div class='center-box'>", unsafe_allow_html=True)
    seed_word = st.text_input("", placeholder="단어 입력", key="init_seed", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)
    
    if seed_word:
        with st.spinner("생각을 그리는 중..."):
            # 시작(루트) 노드 생성
            st.session_state.nodes.append(Node(
                id=seed_word,
                label=seed_word,
                shape="box",
                color={"border": "#8b5cf6", "background": "#ffffff"},
                font={"color": "#333", "size": 16, "bold": True},
                margin={"top": 15, "bottom": 15, "left": 25, "right": 25}
            ))
            expand_graph(seed_word, seed_word)
            st.session_state.is_started = True
            st.rerun()

else:
    # ---------------------------------------------------------
    # [2단계] 마인드맵 캔버스 뷰 (클릭하며 무한 확장)
    # ---------------------------------------------------------
    
    # 💡 3번 사진처럼 완벽한 좌->우 트리 구조를 위한 설정
    config = Config(
        width="100%",
        height="85vh",
        directed=True, 
        physics=True, # 🔥 물리 엔진 ON (노드가 스르륵 날아와 자리 잡는 애니메이션 효과)
        hierarchical=True,
        layout={
            "hierarchical": {
                "enabled": True,
                "direction": "LR",        # 왼쪽에서 오른쪽으로
                "sortMethod": "directed", # 방향성에 맞춰 정렬
                "levelSeparation": 250,   # 좌우 화살표 길이
                "nodeSpacing": 100        # 위아래 상자 간격
            }
        },
        interaction={"hover": True, "dragNodes": False} # 고정된 트리 형태를 위해 드래그 방지
    )
    
    # 그래프 렌더링 (클릭 시 해당 상자의 ID 반환)
    clicked_node_id = agraph(nodes=st.session_state.nodes, edges=st.session_state.edges, config=config)
    
    # 💡 핵심: 클릭한 상자에서 다시 3개의 가지가 스무스하게 뻗어나감
    if clicked_node_id and clicked_node_id not in st.session_state.expanded_nodes:
        # 노드 ID에서 실제 질문 텍스트만 추출
        actual_question = clicked_node_id.split("_||_")[-1] if "_||_" in clicked_node_id else clicked_node_id
        
        with st.spinner("생각을 확장하는 중..."):
            expand_graph(actual_question, clicked_node_id)
            st.rerun()
            
    # 초기화 버튼 (좌측 하단 배치)
    st.markdown("""
        <style>
        .reset-btn { position: fixed; bottom: 20px; left: 20px; z-index: 999; }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("<div class='reset-btn'>", unsafe_allow_html=True)
    if st.button("🔄 처음부터 다시"):
        st.session_state.nodes = []
        st.session_state.edges = []
        st.session_state.expanded_nodes = set()
        st.session_state.is_started = False
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
