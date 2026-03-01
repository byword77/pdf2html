import sys
import os
import json
import logging
import requests
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTabWidget, QDialog, QLineEdit,
    QFormLayout, QComboBox, QDoubleSpinBox, QSpinBox, QTextEdit,
    QMessageBox, QInputDialog
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# 분할된 모듈 임포트
from tab_pdf import PdfTab
from tab_xhtml import XhtmlTab

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "app.log")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
PREPROCESS_PROMPT_FILE = os.path.join(BASE_DIR, "preprocess_prompt.json")

logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", encoding='utf-8')

class ConfigManager:
    def __init__(self):
        self.settings = {
            "api_url": "http://localhost:11434/api/generate",
            "api_key": "",
            "model_name": "", 
            "thumbnail_width": 200,
            "margin_top": 0.0, "margin_bottom": 0.0, "margin_left": 0.0, "margin_right": 0.0,
            "last_prep_prompt": "기본 전처리",
            "extract_style_font": False,
            "extract_style_size": False,
            "extract_style_italic": False,
            "extract_style_bold": False,
            "extract_style_color": False,
            "extract_style_script": False
        }
        self.preprocess_prompts = {
            "기본 전처리": {
                "content": "한국어 맞춤법 규정에 맞춰 불필요하게 줄바꿈 되어 있는 문장을 자연스럽게 연결하고 붙여쓰기를 처리하세요.\n"
                           "**[주의사항]**\n"
                           "1. 제공된 텍스트에 포함된 HTML 태그 및 style 속성은 절대 수정하거나 삭제하지 말고 그대로 유지할 것.\n"
                           "2. 단어 수정이나 문체 변경은 허용하지 않음.\n"
                           "3. 어떠한 부연설명도 추가하지 말고 변환된 HTML 텍스트만 출력할 것.",
                "temperature": 0.1, "top_p": 0.9, "top_k": 40
            }
        }
        self.load_settings()
        self.load_preprocess_prompts()

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    self.settings.update(json.load(f))
            except: pass

    def save_settings(self):
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.settings, f, indent=4, ensure_ascii=False)

    def load_preprocess_prompts(self):
        if os.path.exists(PREPROCESS_PROMPT_FILE):
            try:
                with open(PREPROCESS_PROMPT_FILE, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    for k, v in loaded.items():
                        if isinstance(v, str): self.preprocess_prompts[k] = {"content": v, "temperature": 0.1, "top_p": 0.9, "top_k": 40}
                        else: self.preprocess_prompts[k] = v
            except: pass

    def save_preprocess_prompts(self):
        with open(PREPROCESS_PROMPT_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.preprocess_prompts, f, indent=4, ensure_ascii=False)


class TestGenerationThread(QThread):
    log_signal = pyqtSignal(str)
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings

    def run(self):
        url = self.settings.get("api_url")
        key = self.settings.get("api_key")
        model = self.settings.get("model_name")
        self.log_signal.emit(f"[{datetime.now().strftime('%H:%M:%S')}] 통신 테스트 시작...")
        headers = {"Content-Type": "application/json"}
        if key: headers["Authorization"] = f"Bearer {key}"
        is_ollama = "/api/generate" in url
        prompt = "안녕하세요! 연결 테스트입니다. 짧게 인사만 해주세요."
        payload = {"model": model, "prompt": prompt, "stream": False} if is_ollama else {"model": model, "messages": [{"role": "user", "content": prompt}]}
        
        try:
            res = requests.post(url, headers=headers, json=payload, timeout=120)
            res.raise_for_status()
            reply = res.json().get("response", "") if is_ollama else res.json().get("choices", [{}])[0].get("message", {}).get("content", "")
            self.log_signal.emit(f"[{datetime.now().strftime('%H:%M:%S')}] 성공: {reply}")
        except Exception as e:
            self.log_signal.emit(f"[{datetime.now().strftime('%H:%M:%S')}] 에러: {str(e)}")


class SettingsDialog(QDialog):
    def __init__(self, config_manager, main_window, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.main_window = main_window
        self.setWindowTitle("환경 설정")
        self.resize(650, 500) 
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        tabs = QTabWidget()

        api_tab = QWidget()
        api_layout = QVBoxLayout(api_tab)
        form_layout = QFormLayout()
        self.input_url = QLineEdit()
        self.input_key = QLineEdit()
        self.input_key.setEchoMode(QLineEdit.EchoMode.Password)
        self.combo_model = QComboBox()
        self.combo_model.setEditable(True)
        self.btn_fetch = QPushButton("불러오기")
        
        h_box = QHBoxLayout()
        h_box.addWidget(self.combo_model, 1)
        h_box.addWidget(self.btn_fetch, 0)

        form_layout.addRow("API URL:", self.input_url)
        form_layout.addRow("API Key:", self.input_key)
        form_layout.addRow("모델명:", h_box)
        api_layout.addLayout(form_layout)
        
        self.btn_test = QPushButton("API 연동 확인")
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setFixedHeight(120)
        self.txt_log.setStyleSheet("background-color: #222; color: #00ff00; font-family: Consolas;")
        
        api_layout.addWidget(self.btn_test)
        api_layout.addWidget(self.txt_log)
        tabs.addTab(api_tab, "API 설정")

        prep_tab = QWidget()
        prep_layout = QVBoxLayout(prep_tab)
        self.combo_prep = QComboBox()
        prep_param_layout = QHBoxLayout()
        self.spin_temp = QDoubleSpinBox()
        self.spin_temp.setRange(0.0, 2.0); self.spin_temp.setPrefix("Temp: ")
        self.spin_topp = QDoubleSpinBox()
        self.spin_topp.setRange(0.0, 1.0); self.spin_topp.setPrefix("Top_P: ")
        self.spin_topk = QSpinBox()
        self.spin_topk.setRange(1, 200); self.spin_topk.setPrefix("Top_K: ")
        
        prep_param_layout.addWidget(self.spin_temp)
        prep_param_layout.addWidget(self.spin_topp)
        prep_param_layout.addWidget(self.spin_topk)
        
        self.txt_prep_content = QTextEdit()
        btn_prep_layout = QHBoxLayout()
        self.btn_add = QPushButton("새로 추가")
        self.btn_save = QPushButton("저장")
        self.btn_del = QPushButton("삭제")
        
        for btn in [self.btn_add, self.btn_save, self.btn_del]: btn_prep_layout.addWidget(btn)
        
        prep_layout.addWidget(self.combo_prep)
        prep_layout.addLayout(prep_param_layout) 
        prep_layout.addWidget(self.txt_prep_content)
        prep_layout.addLayout(btn_prep_layout)
        tabs.addTab(prep_tab, "프롬프트 관리")

        layout.addWidget(tabs)
        
        self.btn_fetch.clicked.connect(self.fetch_models)
        self.btn_test.clicked.connect(self.test_connection)
        self.combo_prep.currentTextChanged.connect(self.load_prep)
        self.btn_save.clicked.connect(self.save_prep)
        self.btn_add.clicked.connect(self.add_prep)
        self.btn_del.clicked.connect(self.del_prep)
        
        self.load_from_config()

    def fetch_models(self):
        url = self.input_url.text().strip()
        headers = {"Authorization": f"Bearer {self.input_key.text().strip()}"} if self.input_key.text() else {}
        target_url = url.replace("/api/generate", "/api/tags").replace("/v1/chat/completions", "/v1/models")
        try:
            res = requests.get(target_url, headers=headers, timeout=5)
            if res.status_code == 200:
                data = res.json()
                models = [m["id"] for m in data.get("data", [])] if "data" in data else [m["name"] for m in data.get("models", [])]
                self.combo_model.clear()
                self.combo_model.addItems(models)
                QMessageBox.information(self, "성공", f"{len(models)}개 모델이 감지되었습니다.")
        except Exception as e: QMessageBox.warning(self, "실패", str(e))

    def test_connection(self):
        self.txt_log.clear()
        temp_set = {"api_url": self.input_url.text(), "api_key": self.input_key.text(), "model_name": self.combo_model.currentText()}
        self.test_thread = TestGenerationThread(temp_set, self)
        self.test_thread.log_signal.connect(self.txt_log.append)
        self.test_thread.start()

    def load_from_config(self):
        s = self.config_manager.settings
        self.input_url.setText(s.get("api_url", ""))
        self.input_key.setText(s.get("api_key", ""))
        self.combo_model.setEditText(s.get("model_name", ""))
        self.combo_prep.addItems(self.config_manager.preprocess_prompts.keys())

    def load_prep(self, key):
        if key in self.config_manager.preprocess_prompts:
            data = self.config_manager.preprocess_prompts[key]
            self.txt_prep_content.setText(data.get("content", ""))
            self.spin_temp.setValue(data.get("temperature", 0.1))
            self.spin_topp.setValue(data.get("top_p", 0.9))
            self.spin_topk.setValue(data.get("top_k", 40))

    def save_prep(self):
        key = self.combo_prep.currentText()
        if key:
            self.config_manager.preprocess_prompts[key] = {
                "content": self.txt_prep_content.toPlainText(),
                "temperature": self.spin_temp.value(),
                "top_p": self.spin_topp.value(),
                "top_k": self.spin_topk.value()
            }
            self.config_manager.save_preprocess_prompts()

    def add_prep(self):
        text, ok = QInputDialog.getText(self, "추가", "새 프롬프트 이름:")
        if ok and text:
            self.config_manager.preprocess_prompts[text] = {"content": "", "temperature": 0.1, "top_p": 0.9, "top_k": 40}
            self.combo_prep.addItem(text)
            self.combo_prep.setCurrentText(text)

    def del_prep(self):
        key = self.combo_prep.currentText()
        if key != "기본 전처리" and key in self.config_manager.preprocess_prompts:
            del self.config_manager.preprocess_prompts[key]
            self.combo_prep.removeItem(self.combo_prep.currentIndex())
            self.config_manager.save_preprocess_prompts()

    def closeEvent(self, event): 
        s = self.config_manager.settings
        s["api_url"] = self.input_url.text().strip()
        s["api_key"] = self.input_key.text().strip()
        s["model_name"] = self.combo_model.currentText().strip()
        self.config_manager.save_settings()
        self.main_window.xhtml_tab.update_prompt_combo()
        super().closeEvent(event)


class MainWindow(QMainWindow):
    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.init_ui()
        self.setAcceptDrops(True)

    def init_ui(self):
        self.setWindowTitle("PDF to XHTML & 전처리 툴")
        self.resize(1100, 800)

        self.tabs = QTabWidget()
        self.pdf_tab = PdfTab(self.config_manager, self)
        self.xhtml_tab = XhtmlTab(self.config_manager, self)

        self.tabs.addTab(self.pdf_tab, "1. PDF 변환 (추출)")
        self.tabs.addTab(self.xhtml_tab, "2. XHTML 정리 (LLM)")

        self.btn_settings = QPushButton("⚙️ API/프롬프트 설정")
        self.btn_settings.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_settings.clicked.connect(lambda: SettingsDialog(self.config_manager, self, self).exec())
        
        self.tabs.setCornerWidget(self.btn_settings, Qt.Corner.TopRightCorner)

        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)
        main_layout.setContentsMargins(5, 5, 5, 5)
        
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        self.statusBar().showMessage("준비 완료")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()
        else: super().dragEnterEvent(event)

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if self.tabs.currentIndex() == 0 and path.lower().endswith(".pdf"):
                self.pdf_tab.load_pdf(path)
                event.acceptProposedAction()
            elif self.tabs.currentIndex() == 1 and path.lower().endswith((".xhtml", ".html")):
                self.xhtml_tab.load_xhtml(path)
                event.acceptProposedAction()

    def closeEvent(self, event):
        self.pdf_tab.close_tab()
        self.xhtml_tab.close_tab()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Malgun Gothic", 10))
    config = ConfigManager()
    window = MainWindow(config)
    window.show()
    sys.exit(app.exec())