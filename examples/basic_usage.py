"""
HWP MCP Server 사용 예제

이 예제는 HWP MCP 서버의 기본적인 사용법을 보여줍니다.
실제로 실행하려면 Windows 환경에서 한글 프로그램이 설치되어 있어야 합니다.
"""

# MCP 클라이언트에서 사용하는 방법 (Claude Desktop 등):
#
# 1. 한글 연결
#    hwp_connect visible=true
#
# 2. 새 문서 생성
#    hwp_create
#
# 3. 텍스트 입력
#    hwp_insert_text text="안녕하세요, HWP MCP 서버입니다."
#
# 4. 단락 추가
#    hwp_insert_paragraph
#
# 5. 글꼴 설정
#    hwp_set_font name="맑은 고딕" size=12 bold=true
#
# 6. 텍스트 추가 입력
#    hwp_insert_text text="이 텍스트는 굵은 맑은 고딕으로 작성됩니다."
#
# 7. 표 생성
#    hwp_create_table rows=3 cols=4
#
# 8. 문서 저장
#    hwp_save_as path="C:\\Documents\\test.hwp" format="HWP"
#
# 9. PDF 내보내기
#    hwp_export_pdf path="C:\\Documents\\test.pdf"
#
# 10. 문서 닫기
#     hwp_close


# Python에서 직접 HwpController 사용하기:
if __name__ == "__main__":
    from src.hwp import HwpController

    # Controller 인스턴스 생성
    controller = HwpController()

    try:
        # 한글 연결
        print("Connecting to HWP...")
        controller.connect(visible=True)
        print("Connected!")

        # 새 문서 생성
        print("Creating new document...")
        controller.create_document()

        # 텍스트 입력
        print("Inserting text...")
        controller.insert_text("안녕하세요, HWP MCP 서버 테스트입니다.")
        controller.insert_paragraph()
        controller.insert_text("이것은 두 번째 단락입니다.")

        # 글꼴 설정
        print("Setting font...")
        controller.set_font(name="맑은 고딕", size=14, bold=True)

        # 표 생성
        print("Creating table...")
        controller.insert_paragraph()
        controller.create_table(3, 4)

        # 문서 정보 출력
        info = controller.get_document_info()
        print(f"Document info: {info}")

        print("Done! Check the HWP window.")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # 연결 해제 (문서는 열린 상태로 유지)
        # controller.disconnect()
        pass
