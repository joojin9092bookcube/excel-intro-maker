import io
import json
import google.generativeai as genai
import openpyxl
import streamlit as st

st.set_page_config(
    page_title="론칭작 일러스트 작품 소개서 AI 자동 생성기 (무료)",
    layout="wide",
)

st.title(
    "🎨 [CP작성] 론칭작 일러스트 작품 소개서 AI 자동 생성기 (Gemini 무료 버전)"
)
st.markdown(
    "원고와 인물 이름만 넣으면 Gemini AI가 원고를 분석하여 양식(`guide` 시트 E열"
    " 9~36행) 제약조건에 맞춰 카피 초안을 작성해 줍니다."
)

# 1. 설정 및 입력 영역
st.sidebar.header("⚙️ 기본 설정")
gemini_api_key = st.sidebar.text_input(
    "Google Gemini API Key",
    type="password",
    help="Google AI Studio에서 무료로 발급받으실 수 있습니다.",
)

st.sidebar.subheader("작품 및 인물 설정")
work_title = st.sidebar.text_input("작품 제목", value="빙의자 테스트")
char_1 = st.sidebar.text_input("주인공 1 (예: 류시헌)", value="류시헌")
char_2 = st.sidebar.text_input("주인공 2 (예: 송희민)", value="송희민")

st.subheader("📄 원고 입력")
script_text = st.text_area(
    "분석할 원고 전체 텍스트(또는 시놉시스/초반 분량)를 붙여넣으세요:",
    height=300,
    placeholder=(
        "여기에 원고 텍스트를 입력하면 AI가 분석하여 카피를 추출 및"
        " 생성합니다."
    ),
)

uploaded_file = st.file_uploader(
    "원래 엑셀 양식 파일 업로드 ([CP작성] 론칭작 일러스트 작품 소개서.xlsx)",
    type=["xlsx"],
)


def generate_draft_with_gemini(api_key, title, c1, c2, script):
  genai.configure(api_key=api_key)

  # 최신 환경에서 안정적으로 지원되는 gemini-2.0-flash 모델 지정
  model = genai.GenerativeModel("gemini-2.0-flash")

  prompt = f"""
    당신은 웹소설/웹툰 CP사의 전문 마케터이자 카피라이터입니다.
    아래 전달받은 원고 텍스트를 읽고, 주인공 두 명({c1}, {c2})을 중심으로 론칭작 일러스트 작품 소개서에 들어갈 카피 초안을 작성해 주세요.

    [작품 정보]
    - 작품 제목: {title}
    - 주인공 1: {c1}
    - 주인공 2: {c2}

    [원고 텍스트]
    {script[:8000]}

    [작성 제약 조건]
    1. keywords: 작품을 대표하는 대표 키워드 3개 (#키워드1 #키워드2 #키워드3 형태)
    2. intro: 독자의 흥미를 끌 수 있는 가장 후킹한 인트로 (정확히 2줄, 리스트 형태)
    3. dialogue1_c1: 첫 번째 대화구간 중 {c1}의 대화 (최소 2줄, 최대 3줄, 리스트 형태)
    4. dialogue1_c2: 첫 번째 대화구간 중 {c2}의 대화 (최소 2줄, 최대 3줄, 리스트 형태)
    5. synopsis: 시놉시스를 간략하게 흥미롭게 요약 (정확히 4줄, 리스트 형태)
    6. dialogue2_c2: 두 번째 대화구간 중 {c2}의 대화 (최소 2줄, 최대 3줄, 리스트 형태)
    7. dialogue2_c1: 두 번째 대화구간 중 {c1}의 대화 (최소 2줄, 최대 3줄, 리스트 형태)
    8. plot_summary: 대화에 따른 작품의 전개 내용 흥미롭게 요약 (정확히 3줄, 리스트 형태)
    9. hooking_line: 작품을 읽고 싶어지는 가장 후킹한 대사 (정확히 2줄, 리스트 형태)
    10. outro: 흥미를 유발하는 마무리 소개 (정확히 2줄, 리스트 형태)

    반드시 마크다운 코드 블록(```json 과 ```)을 사용하여 아래 JSON 형식으로만 응답해 주세요. 다른 설명글은 작성하지 마세요.
    ```json
    {{
        "keywords": "#동양풍 #현대판타지 #배틀연애",
        "intro": ["줄1 내용", "줄2 내용"],
        "dialogue1_c1": ["대사1", "대사2", "대사3"],
        "dialogue1_c2": ["대사1", "대사2", "대사3"],
        "synopsis": ["줄1", "줄2", "줄3", "줄4"],
        "dialogue2_c2": ["대사1", "대사2", "대사3"],
        "dialogue2_c1": ["대사1", "대사2", "대사3"],
        "plot_summary": ["줄1", "줄2", "줄3"],
        "hooking_line": ["대사1", "대사2"],
        "outro": ["줄1", "줄2"]
    }}
    ```
    """

  response = model.generate_content(prompt)
  text = response.text.strip()

  # 마크다운 블록 제거 처리
  if "```json" in text:
    text = text.split("```json")[1].split("```")[0].strip()
  elif "```" in text:
    text = text.split("```")[1].split("```")[0].strip()

  return json.loads(text)


if st.button("🚀 AI 초안 분석 및 엑셀 생성하기"):
  if not gemini_api_key:
    st.error("Google Gemini API Key를 입력해주세요.")
  elif not script_text.strip():
    st.error("원고 텍스트를 입력해주세요.")
  elif uploaded_file is None:
    st.error(
        "엑셀 양식 파일([CP작성] 론칭작 일러스트 작품 소개서.xlsx)을"
        " 업로드해주세요."
    )
  else:
    with st.spinner("Gemini AI가 원고를 분석하고 초안을 작성 중입니다..."):
      try:
        ai_data = generate_draft_with_gemini(
            gemini_api_key, work_title, char_1, char_2, script_text
        )

        wb = openpyxl.load_workbook(uploaded_file)
        sheet = wb["guide"]

        sheet["E9"] = ai_data.get("keywords", "")
        sheet["E10"] = "화제작 론칭!"
        sheet["E11"] = work_title

        intro = ai_data.get("intro", [])
        sheet["E12"] = intro[0] if len(intro) > 0 else ""
        sheet["E13"] = intro[1] if len(intro) > 1 else ""

        sheet["D15"] = char_1
        sheet["D18"] = char_2

        d1_c1 = ai_data.get("dialogue1_c1", [])
        d1_c2 = ai_data.get("dialogue1_c2", [])
        for i in range(3):
          sheet[f"E{14+i}"] = d1_c1[i] if i < len(d1_c1) else ""
          sheet[f"E{17+i}"] = d1_c2[i] if i < len(d1_c2) else ""

        syn = ai_data.get("synopsis", [])
        for i in range(4):
          sheet[f"E{20+i}"] = syn[i] if i < len(syn) else ""

        sheet["D25"] = char_2
        sheet["D28"] = char_1

        d2_c2 = ai_data.get("dialogue2_c2", [])
        d2_c1 = ai_data.get("dialogue2_c1", [])
        for i in range(3):
          sheet[f"E{24+i}"] = d2_c2[i] if i < len(d2_c2) else ""
          sheet[f"E{27+i}"] = d2_c1[i] if i < len(d2_c1) else ""

        plot = ai_data.get("plot_summary", [])
        for i in range(3):
          sheet[f"E{30+i}"] = plot[i] if i < len(plot) else ""

        hook = ai_data.get("hooking_line", [])
        sheet["E33"] = hook[0] if len(hook) > 0 else ""
        sheet["E34"] = hook[1] if len(hook) > 1 else ""

        outro = ai_data.get("outro", [])
        sheet["E35"] = outro[0] if len(outro) > 0 else ""
        sheet["E36"] = outro[1] if len(outro) > 1 else ""

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        file_name = f"[CP작성] 론칭작 일러스트 작품 소개서_{work_title}.xlsx"

        st.success("🎉 AI가 초안을 완성했습니다!")
        st.download_button(
            label="📥 완성된 엑셀 파일 다운로드",
            data=output,
            file_name=file_name,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
