# --- 파일 경로 설정 (실제 경로로 변경 필요) ---
# 1번 버튼이 읽을 엑셀 파일 경로 (시스템 설정 확인 버튼에 사용)
EXCEL_FILE_PATH = "C:\\Users\\Public\\Documents\\sample_data.xlsx" 
# PPTX 파일 경로 (K. ECID 메뉴에 사용)
PPTX_FILE_PATH = "C:\\Users\\Public\\Documents\\sample_presentation.pptx"

# --- 서브 메뉴 항목 레이블 매핑 ---
# key: 메뉴 번호 (1~11), value: 메뉴 이름
MENU_LABELS = {
    1: "교환기", 
    2: "CTI", 
    3: "IVR",
    4: "녹취", 
    5: "PDS", 
    6: "ESP-r",
    7: "EUC", 
    8: "EMS", 
    9: "DARS",
    10: "STT", 
    11: "ECID"
}