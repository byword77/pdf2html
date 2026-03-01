import os
import gc
import shutil
import logging
import traceback
import html
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QSplitter, QScrollArea,
    QFileDialog, QMessageBox, QProgressDialog, QDialog,
    QGroupBox, QFormLayout, QDoubleSpinBox, QCheckBox
)
from PyQt6.QtGui import QPixmap, QImage, QColor, QIcon, QPainter, QPen
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QEvent

class ThumbnailWorker(QThread):
    thumb_ready = pyqtSignal(int, str)
    def __init__(self, pdf_path, thumb_dir, page_count, thumb_width, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.thumb_dir = thumb_dir
        self.page_count = page_count
        self.thumb_width = thumb_width
        self._running = True

    def stop(self): self._running = False

    def run(self):
        try:
            doc = fitz.open(self.pdf_path)
            for i in range(self.page_count):
                if not self._running: break
                page = doc.load_page(i)
                try: scale = max(0.1, float(self.thumb_width) / max(1.0, page.rect.width))
                except: scale = 0.2
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                thumb_path = os.path.join(self.thumb_dir, f"page_{i}.png")
                try: png_bytes = pix.tobytes(output="png")
                except: png_bytes = pix.getPNGData() if hasattr(pix, 'getPNGData') else None
                if png_bytes:
                    with open(thumb_path, "wb") as tf: tf.write(png_bytes)
                    self.thumb_ready.emit(i, thumb_path)
                try: del pix
                except: pass
            doc.close()
        except Exception as e:
            logging.error(f"ThumbnailWorker 오류: {e}")

class ExportDocumentThread(QThread):
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, pdf_path, out_path, settings, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.out_path = out_path
        self.settings = settings
        self._running = True

    def stop(self): self._running = False

    def run(self):
        try:
            doc = fitz.open(self.pdf_path)
            total_pages = len(doc)
            content_list = []
            mm_to_pt = 2.83465

            content_list.append(
                '<?xml version="1.0" encoding="UTF-8"?>\n'
                '<!DOCTYPE html>\n'
                '<html xmlns="http://www.w3.org/1999/xhtml">\n'
                '<head>\n'
                '    <title>Extracted Document</title>\n'
                '    <style>\n'
                '        body { font-family: sans-serif; line-height: 1.6; padding: 20px; }\n'
                '    </style>\n'
                '</head>\n'
                '<body>'
            )

            for page_idx in range(total_pages):
                if not self._running: break
                
                page = doc.load_page(page_idx)
                valid_rect = fitz.Rect(
                    self.settings.get("margin_left", 0) * mm_to_pt, 
                    self.settings.get("margin_top", 0) * mm_to_pt, 
                    page.rect.width - self.settings.get("margin_right", 0) * mm_to_pt, 
                    page.rect.height - self.settings.get("margin_bottom", 0) * mm_to_pt
                )

                text_dict = page.get_text("dict", sort=True)
                sizes = []
                for block in text_dict.get("blocks", []):
                    if block.get("type") == 0:  
                        if not fitz.Rect(block.get("bbox")).intersects(valid_rect):
                            continue
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                if span.get("text", "").strip(): sizes.append(round(span.get("size", 10)))
                
                base_size = max(set(sizes), key=sizes.count) if sizes else 10
                page_blocks = []
                
                for block in text_dict.get("blocks", []):
                    if block.get("type") == 0:
                        if not fitz.Rect(block.get("bbox")).intersects(valid_rect):
                            continue

                        block_text = ""
                        for line in block.get("lines", []):
                            line_text = ""
                            max_size_in_line = base_size
                            line_baseline = None
                            
                            for span in line.get("spans", []):
                                span_size = span.get("size", base_size)
                                if span_size > max_size_in_line:
                                    max_size_in_line = span_size
                                    line_baseline = span.get("origin", [0, 0])[1]
                            
                            if line_baseline is None and line.get("spans"):
                                line_baseline = line.get("spans")[0].get("origin", [0, 0])[1]
                            
                            for span in line.get("spans", []):
                                text = span.get("text", "")
                                if not text.strip():
                                    line_text += text
                                    continue
                                
                                size = span.get("size", base_size)
                                flags = span.get("flags", 0)
                                font = span.get("font", "")
                                color_int = span.get("color", 0)
                                
                                hex_color = f"#{color_int:06x}"
                                is_italic = (flags & 2**1) or ("italic" in font.lower())
                                is_bold = (flags & 2**4) or ("bold" in font.lower())
                                
                                is_super = (flags & 1) != 0
                                is_sub = False
                                span_baseline = span.get("origin", [0, 0])[1]
                                
                                if size < max_size_in_line * 0.9:
                                    if span_baseline > line_baseline + 1:
                                        is_sub = True
                                    elif span_baseline < line_baseline - 1:
                                        is_super = True
                                
                                l_space = text[:len(text) - len(text.lstrip())]
                                r_space = text[len(text.rstrip()):]
                                core_text = text.strip()
                                
                                style = ""
                                if self.settings.get("extract_style_font", False):
                                    clean_font = font.split('+')[-1] if '+' in font else font
                                    style += f"font-family:'{clean_font}';"
                                    
                                if self.settings.get("extract_style_size", False):
                                    style += f"font-size:{round(size, 1)}pt;"
                                    
                                if self.settings.get("extract_style_color", False) and hex_color not in ["#000000", "#222222"]:
                                    style += f"color:{hex_color};"
                                    
                                if self.settings.get("extract_style_script", False):
                                    if is_super: style += "vertical-align:super;font-size:0.8em;"
                                    elif is_sub: style += "vertical-align:sub;font-size:0.8em;"
                                
                                core_text = html.escape(core_text)
                                if self.settings.get("extract_style_bold", False) and is_bold:
                                    style += "font-weight:bold;"
                                if self.settings.get("extract_style_italic", False) and is_italic:
                                    style += "font-style:italic;"
                                    
                                if style: core_text = f'<span style="{style}">{core_text}</span>'
                                line_text += f"{l_space}{core_text}{r_space}"
                                
                            tag = "p"
                            if max_size_in_line >= base_size * 2.0: tag = "h1"
                            elif max_size_in_line >= base_size * 1.5: tag = "h2"
                            elif max_size_in_line >= base_size * 1.2: tag = "h3"
                            elif max_size_in_line >= base_size * 1.1: tag = "h4"
                            
                            block_text += f"<{tag}>{line_text.strip()}</{tag}>\n"
                            
                        if block_text.strip(): page_blocks.append(block_text.strip())
                
                if page_blocks:
                    # 전처리 단계 분할용 주석 마커 추가
                    page_blocks.insert(0, f"")
                    content_list.append("\n".join(page_blocks))
                
                self.progress_signal.emit(page_idx + 1, total_pages)
                
            if self._running:
                content_list.append("</body>\n</html>")
                final_text = "\n".join(content_list)

                with open(self.out_path, 'w', encoding='utf-8') as f:
                    f.write(final_text)
                self.finished_signal.emit(self.out_path)
            
            doc.close()
        except Exception as e:
            logging.error(f"저장 오류: {traceback.format_exc()}")
            self.error_signal.emit(str(e))

class ExtractionSettingsDialog(QDialog):
    preview_signal = pyqtSignal(dict)
    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.setWindowTitle("추출 옵션")
        self.resize(320, 380)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        grp_margin = QGroupBox("추출 범위 설정 (여백 제외)")
        grid_margin = QFormLayout()
        self.spin_top = QDoubleSpinBox()
        self.spin_bottom = QDoubleSpinBox()
        self.spin_left = QDoubleSpinBox()
        self.spin_right = QDoubleSpinBox()
        for spin in [self.spin_top, self.spin_bottom, self.spin_left, self.spin_right]:
            spin.setRange(0.0, 300.0)
            spin.setSuffix(" mm")
            spin.setDecimals(1)
            spin.valueChanged.connect(self.emit_preview)

        grid_margin.addRow("상단:", self.spin_top)
        grid_margin.addRow("하단:", self.spin_bottom)
        grid_margin.addRow("좌측:", self.spin_left)
        grid_margin.addRow("우측:", self.spin_right)
        grp_margin.setLayout(grid_margin)

        grp_style = QGroupBox("스타일 추출 옵션")
        style_layout = QVBoxLayout()
        self.chk_font = QCheckBox("글꼴 (Font-family)")
        self.chk_size = QCheckBox("글자 크기 (Font-size)")
        self.chk_italic = QCheckBox("기울임 (Italic)")
        self.chk_bold = QCheckBox("진하게 (Bold)")
        self.chk_color = QCheckBox("글자색 (Color)")
        self.chk_script = QCheckBox("첨자 (위첨자, 아래첨자)")
        
        for chk in [self.chk_font, self.chk_size, self.chk_italic, self.chk_bold, self.chk_color, self.chk_script]:
            style_layout.addWidget(chk)
        grp_style.setLayout(style_layout)

        btn_box = QHBoxLayout()
        btn_ok = QPushButton("확인")
        btn_cancel = QPushButton("취소")
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        btn_box.addStretch()
        btn_box.addWidget(btn_ok)
        btn_box.addWidget(btn_cancel)

        layout.addWidget(grp_margin)
        layout.addWidget(grp_style)
        layout.addLayout(btn_box)
        self.load_data()

    def load_data(self):
        s = self.config_manager.settings
        self.spin_top.setValue(s.get("margin_top", 0.0))
        self.spin_bottom.setValue(s.get("margin_bottom", 0.0))
        self.spin_left.setValue(s.get("margin_left", 0.0))
        self.spin_right.setValue(s.get("margin_right", 0.0))
        
        self.chk_font.setChecked(s.get("extract_style_font", False))
        self.chk_size.setChecked(s.get("extract_style_size", False))
        self.chk_italic.setChecked(s.get("extract_style_italic", False))
        self.chk_bold.setChecked(s.get("extract_style_bold", False))
        self.chk_color.setChecked(s.get("extract_style_color", False))
        self.chk_script.setChecked(s.get("extract_style_script", False))

    def emit_preview(self):
        margins = {
            "margin_top": self.spin_top.value(), 
            "margin_bottom": self.spin_bottom.value(), 
            "margin_left": self.spin_left.value(), 
            "margin_right": self.spin_right.value()
        }
        self.preview_signal.emit(margins)

    def accept(self):
        s = self.config_manager.settings
        s["margin_top"] = self.spin_top.value()
        s["margin_bottom"] = self.spin_bottom.value()
        s["margin_left"] = self.spin_left.value()
        s["margin_right"] = self.spin_right.value()
        
        s["extract_style_font"] = self.chk_font.isChecked()
        s["extract_style_size"] = self.chk_size.isChecked()
        s["extract_style_italic"] = self.chk_italic.isChecked()
        s["extract_style_bold"] = self.chk_bold.isChecked()
        s["extract_style_color"] = self.chk_color.isChecked()
        s["extract_style_script"] = self.chk_script.isChecked()
        
        self.config_manager.save_settings()
        super().accept()

class PdfTab(QWidget):
    def __init__(self, config_manager, main_window):
        super().__init__()
        self.config_manager = config_manager
        self.main_window = main_window
        self.doc = None
        self.current_page_idx = 0
        self.pdf_path = None
        self.zoom_level = 1.0
        self.thumb_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".thumbs")
        self._last_rendered = {"page_idx": None, "pixmap": None}
        self.thumb_worker = None
        self.export_thread = None
        self.preview_margins = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        toolbar = QHBoxLayout()
        self.btn_open = QPushButton("PDF 불러오기")
        self.btn_extract_settings = QPushButton("추출 옵션")
        
        self.btn_save_doc = QPushButton("XHTML 저장") 
        self.btn_reset = QPushButton("초기화") 
        
        self.lbl_info = QLabel("파일 없음")
        self.lbl_info.setStyleSheet("color: gray; padding-left: 10px;")
        
        self.btn_open.clicked.connect(self.open_pdf_dialog)
        self.btn_extract_settings.clicked.connect(self.show_extract_settings)
        self.btn_save_doc.clicked.connect(self.save_document) 
        self.btn_reset.clicked.connect(self.reset_pdf)

        toolbar.addWidget(self.btn_open)
        toolbar.addWidget(self.btn_extract_settings)
        toolbar.addWidget(self.btn_save_doc)
        toolbar.addWidget(self.lbl_info)
        toolbar.addStretch()
        toolbar.addWidget(self.btn_reset)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.thumb_list = QListWidget()
        self.thumb_list.setMinimumWidth(100)
        self.thumb_list.setMaximumWidth(250)
        self.thumb_list.setViewMode(QListWidget.ViewMode.IconMode)
        self.thumb_list.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.thumb_list.setSpacing(10)
        self.thumb_list.setMovement(QListWidget.Movement.Static)
        
        self.thumb_list.setFlow(QListWidget.Flow.TopToBottom)
        self.thumb_list.setWrapping(False)
        
        self.thumb_list.setIconSize(QSize(220, 320))
        self.thumb_list.currentRowChanged.connect(self.change_page)
        self.thumb_list.viewport().installEventFilter(self)

        self.scroll_area = QScrollArea()
        self.pdf_label = QLabel()
        self.pdf_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_area.setWidget(self.pdf_label)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.viewport().installEventFilter(self)

        splitter.addWidget(self.thumb_list)
        splitter.addWidget(self.scroll_area)
        splitter.setSizes([220, 780])

        layout.addLayout(toolbar)
        layout.addWidget(splitter, 1)

    def eventFilter(self, source, event):
        if source == self.thumb_list.viewport() and event.type() == QEvent.Type.Resize:
            width = self.thumb_list.viewport().width()
            target_width = width - 30 
            if target_width > 20: 
                target_height = int(target_width * 1.414)
                self.thumb_list.setIconSize(QSize(target_width, target_height))
                self.thumb_list.setGridSize(QSize(target_width, target_height + 30))
        
        if source == self.scroll_area.viewport() and event.type() == QEvent.Type.Wheel and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            self.zoom_level = max(0.2, min(self.zoom_level * (1.1 if event.angleDelta().y() > 0 else 0.9), 2.0))
            self._last_rendered["page_idx"] = None 
            self.render_main_view()
            if self.main_window.statusBar():
                self.main_window.statusBar().showMessage(f"Zoom: {int(self.zoom_level * 100)}%")
            return True 
        return super().eventFilter(source, event)

    def open_pdf_dialog(self):
        fname, _ = QFileDialog.getOpenFileName(self, "PDF 열기", "", "PDF Files (*.pdf)")
        if fname: self.load_pdf(fname)

    def load_pdf(self, path):
        try:
            try:
                if os.path.exists(self.thumb_dir): shutil.rmtree(self.thumb_dir)
                os.makedirs(self.thumb_dir, exist_ok=True)
            except Exception as e: logging.warning(f"썸네일 디렉토리 초기화 실패: {e}")

            self.doc = fitz.open(path)
            self.pdf_path = path
            self.lbl_info.setText(f"파일: {path}")
            self.lbl_info.setStyleSheet("color: black; padding-left: 10px; font-weight: bold;")
            
            self.thumb_list.clear()
            page_count = len(self.doc)
            placeholder = QPixmap(64, 90)
            placeholder.fill(QColor(240,240,240))
            placeholder_icon = QIcon(placeholder)
            
            for i in range(page_count):
                item = QListWidgetItem()
                item.setIcon(placeholder_icon)
                item.setText(str(i + 1))
                item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
                item.setData(Qt.ItemDataRole.UserRole + 1, None)
                self.thumb_list.addItem(item)
                
            self.thumb_list.setCurrentRow(0)
            if self.main_window.statusBar():
                self.main_window.statusBar().showMessage(f"PDF 로드 완료: {os.path.basename(path)}")

            try:
                if self.thumb_worker and self.thumb_worker.isRunning():
                    self.thumb_worker.stop()
                    self.thumb_worker.wait(0.1)
            except: pass

            thumb_w = 220
            self.thumb_worker = ThumbnailWorker(self.pdf_path, self.thumb_dir, page_count, thumb_w, parent=self)
            self.thumb_worker.thumb_ready.connect(self._on_thumb_ready)
            self.thumb_worker.start()
        except Exception as e: QMessageBox.critical(self, "오류", f"PDF 로드 실패: {e}")

    def _on_thumb_ready(self, page_idx, thumb_path):
        try:
            item = self.thumb_list.item(page_idx)
            if not item: return
            item.setData(Qt.ItemDataRole.UserRole + 1, thumb_path)
            item.setIcon(QIcon(thumb_path))
        except Exception as e: logging.error(f"_on_thumb_ready 오류: {e}")

    def reset_pdf(self):
        if not self.doc: return
        if QMessageBox.question(self, "초기화", "PDF 문서를 닫으시겠습니까?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.close_tab()
            self.lbl_info.setText("파일 없음")
            self.lbl_info.setStyleSheet("color: gray; padding-left: 10px;")
            self.thumb_list.clear()
            self.pdf_label.clear()
            if self.main_window.statusBar():
                self.main_window.statusBar().showMessage("초기화됨")

    def change_page(self, row):
        if self.doc and 0 <= row < len(self.doc):
            self.current_page_idx = row
            self.render_main_view()

    def render_main_view(self):
        if not self.doc: return
        if self._last_rendered.get("page_idx") == self.current_page_idx and self._last_rendered.get("pixmap") is not None:
            self.pdf_label.setPixmap(self._last_rendered["pixmap"])
            return
            
        try:
            if self._last_rendered.get("pixmap") is not None: del self._last_rendered["pixmap"]
        except: pass
        self._last_rendered["page_idx"] = None
        self._last_rendered["pixmap"] = None
        gc.collect()

        page = self.doc.load_page(self.current_page_idx)
        zl = max(0.2, min(self.zoom_level, 2.0))
        pix = page.get_pixmap(matrix=fitz.Matrix(zl, zl))
        qpixmap = QPixmap.fromImage(QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888))

        painter = QPainter(qpixmap)
        pen = QPen(QColor(0, 0, 255))
        pen.setStyle(Qt.PenStyle.DashLine)
        pen.setWidth(2)
        painter.setPen(pen)
        
        s = self.preview_margins if self.preview_margins else self.config_manager.settings
        mm_to_px = 2.83465 * zl
        painter.drawRect(
            int(s.get("margin_left",0)*mm_to_px), 
            int(s.get("margin_top",0)*mm_to_px), 
            int(qpixmap.width()-(s.get("margin_left",0)+s.get("margin_right",0))*mm_to_px), 
            int(qpixmap.height()-(s.get("margin_top",0)+s.get("margin_bottom",0))*mm_to_px)
        )
        painter.end()

        self._last_rendered["page_idx"] = self.current_page_idx
        self._last_rendered["pixmap"] = qpixmap
        self.pdf_label.setPixmap(qpixmap)

        try: del pix; gc.collect()
        except: pass

    def show_extract_settings(self):
        dlg = ExtractionSettingsDialog(self.config_manager, self)
        def update_preview(margins):
            self.preview_margins = margins
            self._last_rendered["page_idx"] = None
            self.render_main_view()
            
        dlg.preview_signal.connect(update_preview)
        dlg.emit_preview()
        
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.preview_margins = None
            self._last_rendered["page_idx"] = None 
            self.render_main_view()
        else:
            self.preview_margins = None
            self._last_rendered["page_idx"] = None
            self.render_main_view()

    def save_document(self):
        if not self.doc or not self.pdf_path: 
            QMessageBox.warning(self, "경고", "PDF 파일을 먼저 로드해주세요.")
            return
            
        default_fname = self.pdf_path.replace(".pdf", "_raw.xhtml")
        fname, _ = QFileDialog.getSaveFileName(self, "XHTML 추출 저장", default_fname, "XHTML Files (*.xhtml);;HTML Files (*.html)")
        if not fname: return
            
        self.progress_dialog = QProgressDialog("XHTML 추출 중...", "취소", 0, len(self.doc), self)
        self.progress_dialog.setWindowTitle("진행 중")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setValue(0)
        self.progress_dialog.canceled.connect(self.cancel_export)
        
        self.export_thread = ExportDocumentThread(self.pdf_path, fname, self.config_manager.settings, self)
        self.export_thread.progress_signal.connect(self.progress_dialog.setValue)
        self.export_thread.finished_signal.connect(self.on_export_finished)
        self.export_thread.error_signal.connect(self.on_export_error)
        
        self.btn_save_doc.setEnabled(False)
        self.btn_open.setEnabled(False)
        self.export_thread.start()

    def cancel_export(self):
        if self.export_thread and self.export_thread.isRunning():
            self.export_thread.stop()
            self.export_thread.wait()
            QMessageBox.information(self, "취소됨", "파일 저장이 취소되었습니다.")
            self.btn_save_doc.setEnabled(True)
            self.btn_open.setEnabled(True)

    def on_export_finished(self, fname):
        self.progress_dialog.close()
        self.btn_save_doc.setEnabled(True)
        self.btn_open.setEnabled(True)
        
        reply = QMessageBox.question(self, "추출 완료", f"XHTML 문서가 추출되었습니다.\n{fname}\n\n지금 바로 LLM 전처리를 진행하시겠습니까?", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.main_window.tabs.setCurrentIndex(1)
            self.main_window.xhtml_tab.load_xhtml(fname)

    def on_export_error(self, err):
        self.progress_dialog.close()
        self.btn_save_doc.setEnabled(True)
        self.btn_open.setEnabled(True)
        QMessageBox.critical(self, "오류", f"문서 저장 중 문제가 발생했습니다:\n{err}")

    def close_tab(self):
        try:
            if self.thumb_worker and self.thumb_worker.isRunning():
                self.thumb_worker.stop()
                self.thumb_worker.wait(0.1)
        except: pass
        try:
            if self.export_thread and self.export_thread.isRunning():
                self.export_thread.stop()
                self.export_thread.wait(0.1)
        except: pass
        try:
            if self.doc: self.doc.close()
        except: pass
        self.doc = None
        try:
            if os.path.exists(self.thumb_dir): shutil.rmtree(self.thumb_dir)
        except: pass
        try:
            if self._last_rendered["pixmap"] is not None: del self._last_rendered["pixmap"]
        except: pass
        self._last_rendered["page_idx"] = None
        self._last_rendered["pixmap"] = None
        gc.collect()