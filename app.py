import tkinter as tk
from tkinter import messagebox, scrolledtext
# 모듈에서 필요한 함수와 상수를 가져옵니다.
from data_handler import read_excel_data, read_pptx_data 
from config import MENU_LABELS

class DataReaderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("내 PC 파일 데이터 뷰어 (모듈화)")
        self.geometry("800x600")
        
        # 메인 프레임: 모든 위젯을 담는 컨테이너 (화면 전환을 위해 사용)
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(expand=True, fill='both')

        # 출력 텍스트 영역 (스크롤 기능 포함)
        self.output_text = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=('Courier New', 10), padx=10, pady=10)
        self.output_text.pack(expand=True, fill='both', padx=20, pady=10)

        # 프로그램 시작 시 메인 메뉴 표시
        self.show_main_menu()

    def update_output(self, content, is_error=False):
        """텍스트 영역의 내용을 업데이트하고 스크롤을 맨 위로 이동시키는 헬퍼 함수"""
        self.output_text.delete(1.0, tk.END) 
        self.output_text.insert(tk.END, content) 
        self.output_text.see(1.0) 
        if is_error:
            messagebox.showerror("오류 발생", content)

    def clear_frame(self):
        """현재 메인 프레임의 모든 위젯을 제거합니다."""
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    # --- 메인 메뉴 화면 구성 ---
    def show_main_menu(self):
        """메인 버튼 2개를 표시하는 화면을 구성합니다."""
        self.clear_frame()

        tk.Label(self.main_frame, text="원하시는 메뉴를 선택해주세요.", 
                 pady=10, font=('Arial', 10, 'bold')).pack()

        button_frame = tk.Frame(self.main_frame)
        button_frame.pack(pady=10)

        # 1번 버튼: 시스템 설정 확인 (Excel 파일 읽기)
        tk.Button(
            button_frame, 
            text="1. 시스템 설정 확인", 
            command=self.handle_excel_click, 
            width=25,
            height=2,
            bg="#2ecc71",
            fg="white"
        ).pack(side=tk.LEFT, padx=10)

        # 2번 버튼: 장애 상황 대응 (서브 메뉴 표시)
        tk.Button(
            button_frame, 
            text="2. 장애 상황 대응", 
            command=self.show_submenu, 
            width=25,
            height=2,
            bg="#3498db",
            fg="white"
        ).pack(side=tk.LEFT, padx=10)
        
        self.update_output("프로그램이 시작되었습니다.\n[1. 시스템 설정 확인]은 Excel 파일 데이터를, [2. 장애 상황 대응]은 서브 메뉴를 보여줍니다.")

    # --- Excel 데이터 처리 클릭 핸들러 ---
    def handle_excel_click(self):
        """Excel 파일 읽기 로직을 data_handler.py로부터 호출합니다."""
        content, path_or_error = read_excel_data()

        if content is None:
            self.update_output(path_or_error, is_error=True)
            return

        output_content = f"--- 1. 시스템 설정 확인 (Excel 파일) 읽기 성공 ---\n"
        output_content += f"파일 경로: {path_or_error}\n\n"
        output_content += content
        self.update_output(output_content)


    # --- 서브 메뉴 (A~K) 화면 구성 ---
    def show_submenu(self):
        """A부터 K까지의 서브 메뉴 버튼을 명시적으로 표시하는 화면을 구성합니다."""
        self.clear_frame()

        tk.Label(self.main_frame, text="[장애 상황 대응] 메뉴를 선택하셨습니다. 상세 항목을 선택해주세요.", 
                 pady=10, font=('Arial', 10, 'bold')).pack()

        submenu_frame = tk.Frame(self.main_frame)
        submenu_frame.pack(pady=20, padx=50)

        button_style = {
            'width': 30,
            'height': 1,
            'bg': "#f39c12", 
            'fg': "white"
        }
        
        # 버튼 텍스트를 config.py의 MENU_LABELS와 직접 매핑하여 생성합니다.
        
        # Row 0
        tk.Button(submenu_frame, text=f"A. {MENU_LABELS[1]}", command=lambda: self.handle_submenu_click(1), **button_style).grid(row=0, column=0, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"B. {MENU_LABELS[2]}", command=lambda: self.handle_submenu_click(2), **button_style).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"C. {MENU_LABELS[3]}", command=lambda: self.handle_submenu_click(3), **button_style).grid(row=0, column=2, padx=10, pady=5)

        # Row 1
        tk.Button(submenu_frame, text=f"D. {MENU_LABELS[4]}", command=lambda: self.handle_submenu_click(4), **button_style).grid(row=1, column=0, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"E. {MENU_LABELS[5]}", command=lambda: self.handle_submenu_click(5), **button_style).grid(row=1, column=1, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"F. {MENU_LABELS[6]}", command=lambda: self.handle_submenu_click(6), **button_style).grid(row=1, column=2, padx=10, pady=5)

        # Row 2
        tk.Button(submenu_frame, text=f"G. {MENU_LABELS[7]}", command=lambda: self.handle_submenu_click(7), **button_style).grid(row=2, column=0, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"H. {MENU_LABELS[8]}", command=lambda: self.handle_submenu_click(8), **button_style).grid(row=2, column=1, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"I. {MENU_LABELS[9]}", command=lambda: self.handle_submenu_click(9), **button_style).grid(row=2, column=2, padx=10, pady=5)
        
        # Row 3
        tk.Button(submenu_frame, text=f"J. {MENU_LABELS[10]}", command=lambda: self.handle_submenu_click(10), **button_style).grid(row=3, column=0, padx=10, pady=5)
        tk.Button(submenu_frame, text=f"K. {MENU_LABELS[11]}", command=lambda: self.handle_submenu_click(11), **button_style).grid(row=3, column=1, padx=10, pady=5)

        # 메인 메뉴로 돌아가기 버튼
        tk.Button(
            self.main_frame, 
            text="⬅️ 메인 메뉴로 돌아가기", 
            command=self.show_main_menu,
            width=20,
            bg="#95a5a6", 
            fg="white"
        ).pack(pady=20)


    def handle_submenu_click(self, menu_num):
        """서브 메뉴 클릭 시 실행되는 함수 (데이터는 data_handler.py에서 처리)"""
        
        # 메뉴 번호(1~11)를 알파벳(A~K)으로 변환
        menu_letter = chr(65 + menu_num - 1)
        item_label = MENU_LABELS.get(menu_num, "알 수 없는 항목")

        # K(11)번 메뉴(ECID)는 PPTX 파일 읽기 로직으로 연결
        if menu_num == 11:
            self.handle_pptx_click(menu_letter, item_label)
            return
            
        # 나머지 메뉴 클릭 시 출력 텍스트 업데이트 (더미 매뉴얼 출력)
        output_content = f"--- {menu_letter}. {item_label} 매뉴얼 항목 선택 ---\n\n"
        output_content += f"**{item_label}** 관련 상세 매뉴얼 내용입니다.\n"
        output_content += "--------------------------------------------------------\n"
        output_content += f"단계 1: {item_label} 시스템의 상태를 확인합니다.\n"
        output_content += f"단계 2: 주요 로그를 검토하고 오류 메시지를 분석합니다.\n"
        output_content += f"단계 3: 표준화된 장애 조치 절차에 따라 대응합니다.\n\n"
        output_content += "이 영역은 실제로는 파일(PDF, TXT) 또는 데이터베이스에서 불러온 상세 정보로 대체됩니다."
        
        self.update_output(output_content)
        
    def handle_pptx_click(self, menu_letter, item_label):
        """PPTX 파일 읽기 로직을 data_handler.py로부터 호출합니다."""
        content, path_or_error = read_pptx_data(menu_letter, item_label)

        if content is None:
            self.update_output(path_or_error, is_error=True)
            return

        # 출력 메시지에 알파벳 메뉴 문자 및 항목 레이블 사용
        output_content = f"--- {menu_letter}. {item_label} 관련 자료 읽기 성공 ---\n"
        output_content += f"파일 경로: {path_or_error}\n\n"
        output_content += content
        self.update_output(output_content)


if __name__ == "__main__":
    app = DataReaderApp()
    app.mainloop()