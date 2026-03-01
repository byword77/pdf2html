import os
import logging
import traceback
import re
import json
import requests
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QComboBox, QTextEdit, QFileDialog, QMessageBox, QProgressDialog
)
from PyQt6.QtGui import QFont, QSyntaxHighlighter, QTextCharFormat, QColor
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QRegularExpression

SESSION = requests.Session()

# ---------------------------------------------------------
# 1. 구문 강조(Syntax Highlighting) 클래스
# ---------------------------------------------------------
class HtmlHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)
        self.highlighting_rules = []

        # 1. HTML 태그 (<p>, <span>, </h1> 등) -> 파란색
        tag_format = QTextCharFormat()
        tag_format.setForeground(QColor("#0055FF"))
        tag_format.setFontWeight(QFont.Weight.Bold)
        self.highlighting_rules.append((QRegularExpression(r'<[^>]*>'), tag_format))

        # 2. 속성명 (class, style 등) -> 자주색
        attr_format = QTextCharFormat()
        attr_format.setForeground(QColor("#AA00AA"))
        self.highlighting_rules.append((QRegularExpression(r'\b[a-zA-Z\-]+\s*(?==)'), attr_format))

        # 3. 속성값 ("텍스트") -> 녹색
        value_format = QTextCharFormat()
        value_format.setForeground(QColor("#00AA00"))
        self.highlighting_rules.append((QRegularExpression(r'"[^"]*"'), value_format))
        self.highlighting_rules.append((QRegularExpression(r"'[^']*'"), value_format))

        # 4. [{[ 단어 ]}] 병합 마커 -> 빨간색 + 노란 배경 강조
        marker_format = QTextCharFormat()
        marker_format.setForeground(QColor("#FF0000"))
        marker_format.setBackground(QColor("#FFFF00"))
        marker_format.setFontWeight(QFont.Weight.Bold)
        self.highlighting_rules.append((QRegularExpression(r'\[\{\[.*?\]\}\]'), marker_format))

        # 5. 문법 오류 태그 <span class="gramma_check"> -> 밑줄 및 빨간색 처리
        error_format = QTextCharFormat()
        error_format.setForeground(QColor("#FF0000"))
        error_format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SpellCheckUnderline)
        error_format.setUnderlineColor(QColor("#FF0000"))
        self.highlighting_rules.append((QRegularExpression(r'<span class="gramma_check">.*?</span>'), error_format))

    def highlightBlock(self, text):
        for pattern, format in self.highlighting_rules:
            iterator = pattern.globalMatch(text)
            while iterator.hasNext():
                match = iterator.next()
                self.setFormat(match.capturedStart(), match.capturedLength(), format)


# ---------------------------------------------------------
# 2. LLM 처리 쓰레드
# ---------------------------------------------------------
class LlmWordCleanThread(QThread):
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, xhtml_path, out_path, prompt_data, settings, parent=None):
        super().__init__(parent)
        self.xhtml_path = xhtml_path
        self.out_path = out_path
        self.prompt_data = prompt_data
        self.settings = settings
        self._running = True

    def stop(self):
        self._running = False

    def call_llm(self, pairs_chunk):
        url = self.settings.get("api_url")
        key = self.settings.get("api_key")
        model_name = self.settings.get("model_name")
        
        user_prompt = self.prompt_data.get("content", "당신은 한국어 맞춤법 및 띄어쓰기 교정 전문가입니다.")
        prompt_text = (
            f"{user_prompt}\n\n"
            "제공되는 단어 쌍(w1, w2)을 분석하여, 띄어쓰기를 하는 것이 맞는지, 붙여쓰기를 하는 것이 맞는지 평가하세요.\n"
            "결과(result) 값은 다음 중 하나만 사용해야 합니다:\n"
            "- 'spaced': 띄어쓰기가 맞는 경우\n"
            "- 'nospaced': 붙여쓰기가 맞는 경우\n"
            "- 'ambiguous': 둘 다 맞거나 둘 다 틀려 판단하기 어려운 경우\n\n"
            "반드시 아래 JSON 배열 포맷으로만 응답하고, 다른 부연 설명이나 단어 수정은 절대 하지 마세요.\n"
        )
        
        temp = self.prompt_data.get("temperature", 0.1)
        top_p = self.prompt_data.get("top_p", 0.9)
        top_k = self.prompt_data.get("top_k", 40)

        input_json = [{"id": idx, "w1": w1, "w2": w2} for idx, (w1, w2) in enumerate(pairs_chunk)]
        full_prompt = prompt_text + "\n[입력]\n" + json.dumps(input_json, ensure_ascii=False)
        
        headers = {"Content-Type": "application/json"}
        if key: headers["Authorization"] = f"Bearer {key}"

        is_ollama = "/api/generate" in url
        
        if is_ollama:
            payload = {
                "model": model_name,
                "prompt": full_prompt,
                "stream": False,
                "format": "json",
                "options": {"temperature": temp, "top_p": top_p, "top_k": top_k}
            }
        else:
            payload = {
                "model": model_name,
                "messages": [
                    {"role": "system", "content": prompt_text},
                    {"role": "user", "content": json.dumps(input_json, ensure_ascii=False)}
                ],
                "temperature": temp,
                "top_p": top_p
            }

        res = SESSION.post(url, headers=headers, json=payload, timeout=600)
        res.raise_for_status()
        data = res.json()
        reply = data.get("response", "") if is_ollama else data.get("choices", [{}])[0].get("message", {}).get("content", "")
        
        # Markdown JSON 블록 제거 처리
        reply = re.sub(r'^```json\s*', '', reply.strip(), flags=re.MULTILINE)
        reply = re.sub(r'\s*```$', '', reply)
        return json.loads(reply)

    def run(self):
        try:
            with open(self.xhtml_path, "r", encoding="utf-8") as f:
                content = f.read()

            pattern = r'\[\{\[(.*?)\s+(.*?)\]\}\]'
            matches = list(re.finditer(pattern, content))
            
            if not matches:
                with open(self.out_path, "w", encoding="utf-8") as f: f.write(content)
                self.finished_signal.emit(self.out_path)
                return

            unique_pairs = list(set((m.group(1), m.group(2)) for m in matches))
            chunk_size = 20
            total_chunks = (len(unique_pairs) + chunk_size - 1) // chunk_size
            results_dict = {}
            
            for i in range(total_chunks):
                if not self._running: return
                chunk = unique_pairs[i*chunk_size:(i+1)*chunk_size]
                self.progress_signal.emit(i+1, total_chunks)
                
                try:
                    llm_resp = self.call_llm(chunk)
                    for item in llm_resp:
                        idx = item.get("id")
                        if idx is not None and 0 <= idx < len(chunk):
                            w1, w2 = chunk[idx]
                            results_dict[(w1, w2)] = item.get("result", "ambiguous")
                except Exception as e:
                    logging.error(f"LLM API 실패: {e}\n{traceback.format_exc()}")
                    for w1, w2 in chunk: results_dict[(w1, w2)] = "ambiguous"

            def replace_func(m):
                w1, w2 = m.group(1), m.group(2)
                res = results_dict.get((w1, w2), "ambiguous")
                if res == "spaced": return f"{w1} {w2}"
                elif res == "nospaced": return f"{w1}{w2}"
                else: return f'<span class="gramma_check">{w1} {w2}</span>'

            final_text = re.sub(pattern, replace_func, content)
            with open(self.out_path, "w", encoding="utf-8") as f: f.write(final_text)
            self.finished_signal.emit(self.out_path)

        except Exception as e:
            logging.error(traceback.format_exc())
            self.error_signal.emit(str(e))


# ---------------------------------------------------------
# 3. XHTML 탭 UI 구성
# ---------------------------------------------------------
class XhtmlTab(QWidget):
    def __init__(self, config_manager, main_window):
        super().__init__()
        self.config_manager = config_manager
        self.main_window = main_window
        self.xhtml_path = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        top = QHBoxLayout()
        self.btn_open = QPushButton("XHTML 불러오기")
        
        self.lbl_option = QLabel("정리 옵션:")
        self.combo_option = QComboBox()
        self.combo_option.addItems(["1. 태그 정리", "2. 단어 정리 (Hunspell)", "3. LLM 정리"])
        self.combo_option.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        
        self.btn_run = QPushButton("실행")

        self.lbl_prompt = QLabel("프롬프트:")
        self.combo_prompt = QComboBox()
        self.combo_prompt.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        
        self.btn_save = QPushButton("변경 사항 저장")

        top.addWidget(self.btn_open)
        top.addWidget(self.lbl_option)
        top.addWidget(self.combo_option)
        top.addWidget(self.btn_run)
        
        top.addStretch() 
        
        top.addWidget(self.lbl_prompt)
        top.addWidget(self.combo_prompt)
        top.addWidget(self.btn_save)

        self.text_preview = QTextEdit()
        self.text_preview.setFont(QFont("Consolas", 10))
        self.highlighter = HtmlHighlighter(self.text_preview.document())

        layout.addLayout(top)
        layout.addWidget(self.text_preview, 1)

        self.btn_open.clicked.connect(self.open_xhtml)
        self.btn_run.clicked.connect(self.run_action)
        self.combo_option.currentTextChanged.connect(self.toggle_prompt_combo)
        self.btn_save.clicked.connect(self.save_changes)

        self.update_prompt_combo()
        self.toggle_prompt_combo(self.combo_option.currentText())

    def update_prompt_combo(self):
        self.combo_prompt.clear()
        self.combo_prompt.addItems(self.config_manager.preprocess_prompts.keys())
        
    def toggle_prompt_combo(self, text):
        is_llm = "LLM" in text
        self.lbl_prompt.setVisible(is_llm)
        self.combo_prompt.setVisible(is_llm)

    def save_changes(self):
        if not self.xhtml_path:
            QMessageBox.warning(self, "경고", "저장할 대상이 없습니다. 파일을 먼저 불러오세요.")
            return
            
        try:
            with open(self.xhtml_path, "w", encoding="utf-8") as f:
                f.write(self.text_preview.toPlainText())
            QMessageBox.information(self, "저장 완료", f"변경 사항이 저장되었습니다:\n{self.xhtml_path}")
        except Exception as e:
            QMessageBox.critical(self, "저장 실패", f"오류가 발생했습니다:\n{str(e)}")

    def open_xhtml(self):
        fname, _ = QFileDialog.getOpenFileName(self, "XHTML 열기", "", "XHTML Files (*.xhtml *.html)")
        if not fname: return
        self.load_xhtml(fname)
        
    def load_xhtml(self, path):
        self.xhtml_path = path
        with open(path, "r", encoding="utf-8") as f:
            self.text_preview.setPlainText(f.read())

    def run_action(self):
        if not self.xhtml_path:
            QMessageBox.warning(self, "경고", "XHTML 파일을 먼저 불러오세요.")
            return

        opt = self.combo_option.currentText()
        if "태그 정리" in opt:
            self.run_tag_merge()
        elif "Hunspell" in opt:
            self.run_hunspell()
        elif "LLM" in opt:
            self.run_llm()

    def run_tag_merge(self):
        with open(self.xhtml_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        lines = content.split('\n')
        i = 0
        while i < len(lines) - 1:
            cur = lines[i].strip()
            
            # 현재 줄이 </p>로 끝나지 않으면 패스
            if not cur.endswith('</p>'):
                i += 1
                continue
                
            # 다음 의미 있는(빈 줄이 아닌) 줄 찾기
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
                
            if j < len(lines):
                nxt = lines[j].strip()
                
                # 찾아낸 다음 줄이 <p>로 시작하는지 검사
                if nxt.startswith('<p>'):
                    m_cur = re.search(r'^(.*)(<span\s+[^>]+>)([^<]*)(</span>\s*</p>)$', cur)
                    m_nxt = re.search(r'^(<p>\s*<span\s+[^>]+>)([^<]*)(</span>)(.*)$', nxt)
                    
                    if m_cur and m_nxt:
                        prefix_cur = m_cur.group(1)
                        cur_tag_open = m_cur.group(2)
                        cur_text = m_cur.group(3)
                        
                        nxt_tag_open = m_nxt.group(1)
                        nxt_text = m_nxt.group(2)
                        suffix_nxt = m_nxt.group(4)
                        
                        attr_cur = cur_tag_open[5:-1].strip()
                        m_nxt_span = re.search(r'<span\s+[^>]+>', nxt_tag_open)
                        
                        if m_nxt_span:
                            attr_nxt = m_nxt_span.group(0)[5:-1].strip()
                                
                            if attr_cur == attr_nxt:
                                clean_cur_text = cur_text.rstrip()
                                
                                if clean_cur_text and clean_cur_text[-1] not in {'.', '?', '!', '"', "'", '‘', '’', '“', '”', '…'}:
                                    m_word_cur = re.search(r'(\S+)(\s*)$', cur_text)
                                    m_word_nxt = re.search(r'^(\s*)(\S+)', nxt_text)
                                    
                                    if m_word_cur and m_word_nxt:
                                        last_word = m_word_cur.group(1)
                                        first_word = m_word_nxt.group(2)
                                        
                                        new_cur_text = cur_text[:m_word_cur.start(1)]
                                        new_nxt_text = nxt_text[m_word_nxt.end(2):]
                                        
                                        merged_text = f"{new_cur_text}[{{[{last_word} {first_word}]}}]{new_nxt_text}"
                                        
                                        new_combined_line = f"{prefix_cur}{cur_tag_open}{merged_text}</span>{suffix_nxt}"
                                        
                                        # 줄 내용 교체 후, 중간에 있던 빈 줄을 포함해 원본 다음 줄(들)을 삭제
                                        lines[i] = new_combined_line
                                        del lines[i+1:j+1]
                                        continue # 삭제 후, 새로 갱신된 i번째 줄을 기준으로 다음 줄과 또 병합될 수 있으므로 다시 반복
            i += 1
            
        out_path = self.xhtml_path.replace(".xhtml", "_merged.xhtml")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write('\n'.join(lines))
            
        self.xhtml_path = out_path
        self.text_preview.setPlainText('\n'.join(lines))
        QMessageBox.information(self, "완료", "태그 병합 정리가 완료되었습니다.")

    def run_hunspell(self):
        try:
            from spylls.hunspell import Dictionary
            
            base_dir = os.path.dirname(os.path.abspath(__file__))
            dict_base_path = os.path.join(base_dir, "hunspell", "ko")
            
            if not os.path.exists(dict_base_path + ".dic") or not os.path.exists(dict_base_path + ".aff"):
                QMessageBox.warning(self, "오류", f"hunspell 폴더 내에 사전 파일이 없습니다.\n확인 경로:\n{dict_base_path}.dic / .aff")
                return

            hobj = Dictionary.from_files(dict_base_path)
            
        except ImportError:
            QMessageBox.warning(self, "오류", "spylls 라이브러리가 설치되지 않았습니다.\n(명령어: pip install spylls)")
            return
        except Exception as e:
            QMessageBox.warning(self, "오류", f"사전 로드 중 문제가 발생했습니다:\n{str(e)}")
            return

        with open(self.xhtml_path, "r", encoding="utf-8") as f: 
            content = f.read()
            
        def process_hunspell(m):
            w1, w2 = m.group(1), m.group(2)
            
            ok_spaced = hobj.lookup(w1) and hobj.lookup(w2)
            ok_nospaced = hobj.lookup(w1 + w2)
            
            if ok_spaced and not ok_nospaced: 
                return f"{w1} {w2}"
            elif not ok_spaced and ok_nospaced: 
                return f"{w1}{w2}"
            else: 
                return f'<span class="gramma_check">{w1} {w2}</span>'
                
        final_text = re.sub(r'\[\{\[(.*?)\s+(.*?)\]\}\]', process_hunspell, content)
        
        out_path = self.xhtml_path.replace(".xhtml", "_hunspell.xhtml")
        with open(out_path, "w", encoding="utf-8") as f: 
            f.write(final_text)
            
        self.xhtml_path = out_path
        self.text_preview.setPlainText(final_text)
        QMessageBox.information(self, "완료", "Hunspell 단어 정리가 완료되었습니다.")

    def run_llm(self):
        key = self.combo_prompt.currentText()
        prompt_data = self.config_manager.preprocess_prompts.get(key, {})
        
        out_path = self.xhtml_path.replace(".xhtml", "_llm_clean.xhtml")
        self.progress = QProgressDialog("LLM 단어 교정 중...", "취소", 0, 100, self)
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)

        self.thread = LlmWordCleanThread(self.xhtml_path, out_path, prompt_data, self.config_manager.settings, self)
        self.thread.progress_signal.connect(self.progress.setValue)
        self.thread.finished_signal.connect(self.on_llm_finished)
        self.thread.error_signal.connect(self.on_error)
        self.progress.canceled.connect(self.thread.stop)
        
        self.thread.start()

    def on_llm_finished(self, path):
        self.progress.close()
        self.xhtml_path = path
        with open(path, "r", encoding="utf-8") as f:
            self.text_preview.setPlainText(f.read())
        QMessageBox.information(self, "완료", f"LLM 교정 완료:\n{path}")

    def on_error(self, msg):
        self.progress.close()
        QMessageBox.critical(self, "오류", msg)

    def close_tab(self):
        try:
            if hasattr(self, "thread") and self.thread.isRunning():
                self.thread.stop()
                self.thread.wait(100)
        except: pass