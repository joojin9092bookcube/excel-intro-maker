import io
import openpyxl
import streamlit as st

st.set_page_config(
    page_title="론칭작 일러스트 작품 소개서 초안 생성기", layout="wide"
)

st.title("🎨 [CP작성] 론칭작 일러스트 작품 소개서 자동 생성기")
st.markdown(
    "지정된 양식(`guide` 시트 E열 9~36행)에 맞춰 원고와 인물 설정을 기반으로 초안을 작성하고 엑셀 파일로 다운로드합니다."
)

# 1. 설정 페이지 (인물 설정 및 원고 입력)
st.sidebar.header("⚙️ 작품 및 인물 설정")
work_title = st.sidebar.text_input("작품 제목", value="혼불")

st.sidebar.subheader("인물 설정 (2명)")
char_1 = st.sidebar.text_input("인물 1 이름 (예: 윤태희)", value="윤태희")
char_2 = st.sidebar.text_input("인물 2 이름 (예: 김재겸)", value="김재겸")

st.sidebar.subheader("원고 및 키워드 입력")
keyword_input = st.sidebar.text_input(
    "메인 키워드 (띄어쓰기 구분, # 자동생성)",
    value="동양풍 현대판타지 배틀연애",
)
intro_text = st.sidebar.text_area(
    "인트로 2줄 (독자의 흥미를 끌 가장 후킹한 2줄)",
    value="귀시고 인간이고 다 싫어.\n세상을 등진 외톨이.",
)
dialogue_1 = st.sidebar.text_area(
    "첫 번째 대화 구간 (인물1, 인물2 대화 3줄씩)",
    value="멋대로 과거를 엿봐서 기분 나빴어?\n알고 있겠지만,\n내가 너한테 관심이 많거든.\n---\n무슨 배짱으로 낯짝을\n들이밀었나 했더니...\n개수작 부리지 마.",
)
synopsis_text = st.sidebar.text_area(
    "시놉시스 요약 (4줄 고정)",
    value="귀신이고 인간이고 다 싫어.\n세상을 등진 외톨이,\n'볼 줄 아는' 재겸을 찾아\n귀신 잡는 나례청에서 온 남자, 윤태희.",
)
dialogue_2 = st.sidebar.text_area(
    "두 번째 대화 구간 (인물2, 인물1 대화 3줄씩)",
    value="봐주니까 정도를 모르고 기어오르지.\n죽었다 깨어나도,\n나자가 될 일은 없어.\n---\n왜?\n...나자를\n잘 아나 봐?",
)
plot_summary = st.sidebar.text_area(
    "전개 내용 요약 (3줄 고정)",
    value=(
        "어릴 적 만났던 나자 스승에게 배신 당해\n죽고자 해도 죽을 수 없는"
        " 저주에 걸린 재겸은\n자신에게 접근하는 태희를 거부하는데…."
    ),
)
hooking_line = st.sidebar.text_area(
    "가장 후킹한 대사 (2줄 고정)",
    value="어디부터 잘못되었냐고?\n나자를 만난 때부터야.",
)
outro_text = st.sidebar.text_area(
    "마무리 소개 (2줄 고정)",
    value=(
        "“죽었다 깨어나도 나자가 될 일은 없어.”\n'볼 줄 아는' 소년이 선언한"
        " 선전포고의 끝은?"
    ),
)

uploaded_file = st.file_uploader(
    "원래 엑셀 양식 파일 업로드 ([CP작성] 론칭작 일러스트 작품 소개서.xlsx)",
    type=["xlsx"],
)

if st.button("🚀 초안 엑셀 파일 생성하기"):
  if uploaded_file is None:
    st.error("기준이 되는 엑셀 양식 파일을 먼저 업로드해주세요!")
  else:
    try:
      # 엑셀 파일 로드
      wb = openpyxl.load_workbook(uploaded_file)
      if "guide" not in wb.sheetnames:
        st.error("엑셀 파일에 'guide' 시트가 존재하지 않습니다.")
      else:
        sheet = wb["guide"]

        # 키워드 처리 (# 붙이기)
        keywords = [
            f"#{k.strip()}" if not k.startswith("#") else k.strip()
            for k in keyword_input.split()
        ]
        kw_str = " ".join(keywords[:3])

        # 인물 대화 분할 처리 함수
        def parse_dialogue(text):
          lines = [
              line.strip() for line in text.split("\n") if line.strip()
          ]
          return lines

        d1_lines = parse_dialogue(dialogue_1)
        d2_lines = parse_dialogue(dialogue_2)
        intro_lines = [
            l.strip() for l in intro_text.split("\n") if l.strip()
        ][:2]
        syn_lines = [
            l.strip() for l in synopsis_text.split("\n") if l.strip()
        ][:4]
        plot_lines = [
            l.strip() for l in plot_summary.split("\n") if l.strip()
        ][:3]
        hook_lines = [
            l.strip() for l in hooking_line.split("\n") if l.strip()
        ][:2]
        out_lines = [l.strip() for l in outro_text.split("\n") if l.strip()][:2]

        # E열 9~36행 매핑 채우기
        # E9: 메인카피 (키워드)
        sheet["E9"] = kw_str
        sheet["E10"] = "톨쥬 화제작 리다무 론칭!"  # 후킹 강조

        # E11: 작품 제목
        sheet["E11"] = work_title

        # E12~E13: 인트로 (2줄)
        sheet["E12"] = intro_lines[0] if len(intro_lines) > 0 else ""
        sheet["E13"] = intro_lines[1] if len(intro_lines) > 1 else ""

        # D15, D18: 인물 이름 매핑
        sheet["D15"] = char_1
        sheet["D18"] = char_2

        # 첫 번째 대화 (14~16행)
        d1_part1 = d1_lines[:3] if len(d1_lines) >= 3 else d1_lines + [""] * (3 - len(d1_lines))
        sheet["E14"] = d1_part1[0] if len(d1_part1) > 0 else ""
        sheet["E15"] = d1_part1[1] if len(d1_part1) > 1 else ""
        sheet["E16"] = d1_part1[2] if len(d1_part1) > 2 else ""

        # 첫 번째 대화 (17~19행)
        d1_part2 = d1_lines[3:6] if len(d1_lines) >= 6 else d1_lines[3:] + [""] * (6 - len(d1_lines))
        sheet["E17"] = d1_part2[0] if len(d1_part2) > 0 else ""
        sheet["E18"] = d1_part2[1] if len(d1_part2) > 1 else ""
        sheet["E19"] = d1_part2[2] if len(d1_part2) > 2 else ""

        # 시놉시스 요약 (20~23행)
        for idx, row in enumerate(range(20, 24)):
          sheet[f"E{row}"] = (
              syn_lines[idx] if idx < len(syn_lines) else ""
          )

        # D25, D28: 인물 이름 교차 매핑
        sheet["D25"] = char_2
        sheet["D28"] = char_1

        # 두 번째 대화 (24~26행)
        d2_part1 = d2_lines[:3] if len(d2_lines) >= 3 else d2_lines + [""] * (3 - len(d2_lines))
        sheet["E24"] = d2_part1[0] if len(d2_part1) > 0 else ""
        sheet["E25"] = d2_part1[1] if len(d2_part1) > 1 else ""
        sheet["E26"] = d2_part1[2] if len(d2_part1) > 2 else ""

        # 두 번째 대화 (27~29행)
        d2_part2 = d2_lines[3:6] if len(d2_lines) >= 6 else d2_lines[3:] + [""] * (6 - len(d2_lines))
        sheet["E27"] = d2_part2[0] if len(d2_part2) > 0 else ""
        sheet["E28"] = d2_part2[1] if len(d2_part2) > 1 else ""
        sheet["E29"] = d2_part2[2] if len(d2_part2) > 2 else ""

        # 전개 내용 요약 (30~32행)
        for idx, row in enumerate(range(30, 33)):
          sheet[f"E{row}"] = (
              plot_lines[idx] if idx < len(plot_lines) else ""
          )

        # 가장 후킹한 대사 (33~34행)
        sheet["E33"] = hook_lines[0] if len(hook_lines) > 0 else ""
        sheet["E34"] = hook_lines[1] if len(hook_lines) > 1 else ""

        # 마무리 소개 (35~36행)
        sheet["E35"] = out_lines[0] if len(out_lines) > 0 else ""
        sheet["E36"] = out_lines[1] if len(out_lines) > 1 else ""

        # 결과 파일 바이트로 변환
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        file_name = f"[CP작성] 론칭작 일러스트 작품 소개서_{work_title}.xlsx"

        st.success(
            "🎉 초안 작성이 성공적으로 완료되었습니다! 아래 버튼을 눌러"
            " 다운로드하세요."
        )
        st.download_button(
            label="📥 완성된 엑셀 파일 다운로드",
            data=output,
            file_name=file_name,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml"
                ".spreadsheetml.sheet"
            ),
        )
    except Exception as e:
      st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
