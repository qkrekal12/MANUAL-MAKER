# Streamlit 매뉴얼 스튜디오

기존 웹앱과 독립적으로 실행하는 Streamlit 버전입니다. `run_streamlit.py`는 Streamlit Cloud와 `streamlit run` 명령에서 사용하는 앱 진입 파일입니다. 로컬에서는 `run_streamlit.bat` 또는 `python launch_local.py`로 실행할 수 있습니다. 최초 로컬 실행에는 인터넷 연결이 필요합니다.

## Streamlit Community Cloud 배포

GitHub 저장소의 앱 루트에 `run_streamlit.py`, `app.py`, `requirements.txt` 파일을 함께 둡니다. Streamlit 배포 설정에서 **Main file path**를 `run_streamlit.py`로 지정합니다. OpenAI API 키는 앱의 **Settings → Secrets**에 `OPENAI_API_KEY = "본인_API_키"` 형식으로 등록하세요. 키를 GitHub 저장소에 올리면 안 됩니다.

Cloud는 `requirements.txt`를 설치하고 `run_streamlit.py`를 앱으로 실행합니다. 이 파일은 가상환경이나 `pip install`을 실행하지 않으므로 배포 서버에서 중첩 설치를 시도하지 않습니다.

## 사용 방법

1. `run_streamlit.bat`를 실행합니다.
2. 기능명·서비스명과 개발 메모를 입력하고 이번 매뉴얼에 사용할 화면 이미지를 첨부합니다.
3. 여러 이미지는 파일명에 포함된 숫자를 기준으로 정렬됩니다.
4. 이미지의 버튼/입력 위치를 클릭해 번호를 추가하고 오른쪽 입력란에 항목명과 설명을 적습니다. 필요하지 않으면 번호 설명 없이도 진행할 수 있습니다.
5. OpenAI API 키를 환경 변수 `OPENAI_API_KEY`, 프로젝트의 `.env`, 또는 Cloud의 Streamlit Secrets에 설정합니다. 앱 화면에 임시 입력할 수도 있으며 저장되지 않습니다.
6. **화면 분석하고 매뉴얼 작성하기**를 누르고 결과를 검토한 다음 HTML, Word 또는 PDF를 다운로드합니다.

기본 모델은 `gpt-6-luna`입니다. 필요하면 `OPENAI_MODEL` 환경 변수 또는 화면의 모델 입력란에서 바꿀 수 있습니다. API 키는 출력하거나 별도 파일로 복사하지 않습니다. 이미지는 매뉴얼 작성 버튼을 누를 때 OpenAI API로 전송되며 이 앱은 업로드 파일을 디스크에 저장하지 않습니다.

앱은 현재 PC에서만 접속하도록 `127.0.0.1:8767`에 바인딩합니다. 종료할 때는 실행된 명령 창을 닫으세요.
