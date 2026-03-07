import sys
import os
import shutil
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QPushButton, QVBoxLayout, QHBoxLayout, QFormLayout,
    QFileDialog, QComboBox, QMessageBox, QProgressDialog, 
    QSpinBox, QCheckBox, QTabWidget, QGroupBox, QSizePolicy
)
from PySide6.QtGui import QPixmap, QIcon, QDesktopServices, QPainter, QPen
from PySide6.QtCore import Qt, QUrl, QSize, Signal, QPoint, QRect, QSettings

from PIL import Image
from PIL.ImageQt import ImageQt
from qt_material import apply_stylesheet

from core import gerar_certificado as gerar_imagem_certificado
from updater import check_for_update, download_and_update, get_current_version

class PreviewLabel(QLabel):
    signature_moved = Signal(int, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setAlignment(Qt.AlignCenter)
        self.signature_pixmap = None
        self.signature_rect = QRect()
        self.full_image_size = QSize()
        self.is_dragging = False
        self.drag_offset = QPoint()
        self.is_hovering = False
        self.setStyleSheet("background-color: #2b2b2b; border-radius: 8px;")

    def set_base_pixmap(self, pixmap, original_size):
        self.setPixmap(pixmap)
        self.full_image_size = QSize(original_size[0], original_size[1])

    def update_signature(self, signature_path, pos_x, pos_y, width):
        if signature_path and os.path.exists(signature_path):
            self.signature_pixmap = QPixmap(signature_path)
            self._update_signature_rect(pos_x, pos_y, width)
        else:
            self.signature_pixmap = None
        self.update()

    def clear_signature(self):
        self.signature_pixmap = None
        self.update()

    def _get_displayed_pixmap_rect(self):
        if not self.pixmap() or self.pixmap().isNull() or not self.full_image_size.isValid():
            return QRect()
        scaled_size = self.pixmap().size().scaled(self.size(), Qt.KeepAspectRatio)
        x = (self.width() - scaled_size.width()) / 2
        y = (self.height() - scaled_size.height()) / 2
        return QRect(QPoint(int(x), int(y)), scaled_size)

    def _original_coords_to_widget_rect(self, ox, oy, o_width):
        if not self.full_image_size.isValid(): return QRect()
        displayed_rect = self._get_displayed_pixmap_rect()
        scale_ratio = displayed_rect.width() / self.full_image_size.width()
        sig_orig_ratio = self.signature_pixmap.height() / self.signature_pixmap.width()
        o_height = o_width * sig_orig_ratio
        wx = displayed_rect.x() + (ox * scale_ratio)
        wy = displayed_rect.y() + (oy * scale_ratio)
        return QRect(int(wx), int(wy), int(o_width * scale_ratio), int(o_height * scale_ratio))

    def _widget_pos_to_original_coords(self, w_pos):
        if not self.full_image_size.isValid(): return QPoint(0, 0)
        displayed_rect = self._get_displayed_pixmap_rect()
        if not displayed_rect.contains(w_pos): return None
        scale_ratio = self.full_image_size.width() / displayed_rect.width()
        relative_pos = w_pos - displayed_rect.topLeft()
        return QPoint(int(relative_pos.x() * scale_ratio), int(relative_pos.y() * scale_ratio))

    def _update_signature_rect(self, pos_x, pos_y, width):
        if self.signature_pixmap:
            self.signature_rect = self._original_coords_to_widget_rect(pos_x, pos_y, width)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.signature_pixmap and self.signature_rect.contains(event.pos()):
            self.is_dragging = True
            self.drag_offset = event.pos() - self.signature_rect.topLeft()
            self.setCursor(Qt.ClosedHandCursor)

    def mouseMoveEvent(self, event):
        if self.is_dragging:
            new_top_left = event.pos() - self.drag_offset
            original_pos = self._widget_pos_to_original_coords(new_top_left)
            if original_pos:
                self.signature_moved.emit(original_pos.x(), original_pos.y())
        elif self.signature_pixmap and self.signature_rect.contains(event.pos()):
            if not self.is_hovering:
                self.is_hovering = True
                self.setCursor(Qt.OpenHandCursor)
        elif self.is_hovering:
            self.is_hovering = False
            self.setCursor(Qt.ArrowCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.is_dragging = False
            self.setCursor(Qt.OpenHandCursor if self.is_hovering else Qt.ArrowCursor)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.signature_pixmap and not self.signature_pixmap.isNull():
            painter = QPainter(self)
            painter.drawPixmap(self.signature_rect, self.signature_pixmap)
            pen = QPen(Qt.red, 2, Qt.DashLine)
            painter.setPen(pen)
            painter.drawRect(self.signature_rect)
            painter.end()

def get_asset_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.settings = QSettings("CAHIS", "GeradorCertificados")
        
        self.signature_path = None 
        self.setup_user_directories()
        
        self.setWindowTitle(f"Gerador de Certificados - {get_current_version()}")
        self.setWindowIcon(QIcon("icone.ico"))
        self.setMinimumSize(1200, 800)
        
        self.funcoes_horas = { 'Ouvinte': 5, 'Palestrante': 2, 'Apresentador(a)': 10, 'Organizador(a)': 10, 'Mediador(a)': 2, 'Debatedor(a)': 2, 'Outro': '' }
        self.tipos_de_atividade = [ "Evento", "Seminário", "Palestra", "Mesa redonda", "Apresentação de trabalho", "Curso", "Oficina", "Projeto de extensão", "Evento científico", "Disciplina não curricular", "Atividade Institucionalizada", "Estágio extracurricular", "Curso de língua estrangeira", "Concurso de monografia", "Bolsa de Iniciação Científica", "Competição esportiva", "Outro" ]
        self.atividades_sem_nome = ["Disciplina não curricular", "Bolsa de Iniciação Científica"]

        self.templates = self.load_resources(self.models_dir, ('.png', '.jpg', '.jpeg'))
        self.template_path = os.path.join(self.models_dir, self.templates[0]) if self.templates else ""
        self.fonts = self.load_fonts()
        self.font_paths = list(self.fonts.values())[0] if self.fonts else {}
        
        self.setStyleSheet("""
            QGroupBox { font-weight: bold; border: 1px solid #555; border-radius: 8px; margin-top: 10px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top center; padding: 0 5px; }
            QPushButton { border-radius: 6px; padding: 8px; font-weight: 500; }
            QPushButton#btn_social { background-color: transparent; border: none; }
            QPushButton#btn_social:hover { background-color: rgba(255, 255, 255, 0.1); }
            QPushButton#btn_gerar { background-color: #2ecc71; color: white; font-weight: bold; font-size: 14px; padding: 12px; }
            QPushButton#btn_gerar:hover { background-color: #27ae60; }
            QPushButton#btn_lote { background-color: #f1c40f; color: #2c3e50; font-weight: bold; }
            QLabel#preview_placeholder { color: #888; font-style: italic; }
        """)

        self.init_ui()
        self.on_atividade_change(self.atividade_combo.currentText())
        self.on_funcao_change(self.funcao_combo.currentText())
        self.toggle_date_field()
        self.update_preview()
        
        self.check_updates()

    def check_updates(self):
        allow_beta = self.beta_checkbox.isChecked()
        latest, url, changelog = check_for_update(allow_beta)
        
        if latest:
            msg = QMessageBox(self)
            msg.setWindowTitle('Atualização Disponível')
            msg.setText(f"A versão {latest} está disponível. Deseja atualizar agora?\n\nO programa será reiniciado.")
            msg.setDetailedText(f"O que há de novo:\n\n{changelog}")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            
            if msg.exec() == QMessageBox.Yes:
                QMessageBox.information(self, "Baixando", "O download começou. Aguarde o reinício automático.")
                download_and_update(url)

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        header_layout = QHBoxLayout()
        title_label = QLabel("Gerador de Certificados")
        title_label.setStyleSheet("font-size: 26px; font-weight: bold; color: #ffffff;")
        
        self.twitter_button = QPushButton(); self.twitter_button.setObjectName("btn_social")
        self.twitter_button.setIcon(QIcon(os.path.join(self.icons_dir, "twitter_icon.png")))
        self.instagram_button = QPushButton(); self.instagram_button.setObjectName("btn_social")
        self.instagram_button.setIcon(QIcon(os.path.join(self.icons_dir, "instagram_icon.png")))
        self.bug_button = QPushButton(); self.bug_button.setObjectName("btn_social")
        self.bug_button.setIcon(QIcon(os.path.join(self.icons_dir, "bug_icon.png")))
        
        for btn in [self.twitter_button, self.instagram_button, self.bug_button]:
            btn.setIconSize(QSize(24, 24))
            btn.setFixedSize(36, 36)

        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.twitter_button)
        header_layout.addWidget(self.instagram_button)
        header_layout.addWidget(self.bug_button)
        
        main_layout.addLayout(header_layout)

        content_layout = QHBoxLayout()
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("QTabWidget::pane { border: 1px solid #444; border-radius: 6px; }")
        
        tab_part = QWidget()
        form_part = QFormLayout(tab_part)
        form_part.setSpacing(15)
        
        self.pessoa_input = QLineEdit()
        self.funcao_combo = QComboBox(); self.funcao_combo.addItems(self.funcoes_horas.keys())
        self.funcao_custom_input = QLineEdit(); self.funcao_custom_input.setVisible(False); self.funcao_custom_input.setPlaceholderText("Especifique a função")
        self.doc_tipo_combo = QComboBox(); self.doc_tipo_combo.addItems(['Nenhum', 'CPF', 'Matrícula'])
        self.doc_input = QLineEdit()
        self.trabalho_label = QLabel("Título do Trabalho:"); self.trabalho_input = QLineEdit()
        
        form_part.addRow("Nome Completo:", self.pessoa_input)
        form_part.addRow("Função:", self.funcao_combo)
        form_part.addRow("", self.funcao_custom_input)
        form_part.addRow("Tipo Documento:", self.doc_tipo_combo)
        form_part.addRow("Nº Documento:", self.doc_input)
        form_part.addRow(self.trabalho_label, self.trabalho_input)
        
        tab_evt = QWidget()
        form_evt = QFormLayout(tab_evt)
        form_evt.setSpacing(15)
        
        self.atividade_combo = QComboBox(); self.atividade_combo.addItems(self.tipos_de_atividade)
        self.evento_input = QLineEdit()
        self.horas_input = QLineEdit()
        self.date_checkbox = QCheckBox("Incluir Data/Período")
        self.date_inicio_label = QLabel("Data (Início):")
        self.date_inicio_input = QLineEdit(); self.date_inicio_input.setPlaceholderText("Ex: 23/03/2026")
        self.date_fim_label = QLabel("Data (Fim):")
        self.date_fim_input = QLineEdit(); self.date_fim_input.setPlaceholderText("Ex: 27/03/2026 (Opcional)")
        
        form_evt.addRow("Tipo de Atividade:", self.atividade_combo)
        form_evt.addRow("Nome do Evento:", self.evento_input)
        form_evt.addRow("Carga Horária:", self.horas_input)
        form_evt.addRow(self.date_checkbox)
        form_evt.addRow(self.date_inicio_label, self.date_inicio_input)
        form_evt.addRow(self.date_fim_label, self.date_fim_input)

        tab_config = QWidget()
        layout_config = QVBoxLayout(tab_config)
        
        grp_visual = QGroupBox("Aparência")
        form_visual = QFormLayout(grp_visual)
        self.template_combo = QComboBox(); self.template_combo.addItems(self.templates) if self.templates else None
        self.font_combo = QComboBox(); self.font_combo.addItems(self.fonts.keys()) if self.fonts else None
        self.font_size_input = QSpinBox(); self.font_size_input.setRange(10, 200); self.font_size_input.setValue(50); self.font_size_input.setSuffix(" pt")
        self.pos_x_input = QSpinBox(); self.pos_x_input.setRange(0, 4000); self.pos_x_input.setValue(250)
        self.pos_y_input = QSpinBox(); self.pos_y_input.setRange(0, 4000); self.pos_y_input.setValue(600)
        self.italic_checkbox = QCheckBox("Itálico no nome do evento")
        
        form_visual.addRow("Modelo:", self.template_combo)
        form_visual.addRow("Fonte:", self.font_combo)
        form_visual.addRow("Tamanho:", self.font_size_input)
        form_visual.addRow("Posição X:", self.pos_x_input)
        form_visual.addRow("Posição Y:", self.pos_y_input)
        form_visual.addRow(self.italic_checkbox)
        
        grp_sign = QGroupBox("Assinatura Digital")
        layout_sign = QVBoxLayout(grp_sign)
        self.signature_checkbox = QCheckBox("Habilitar Assinatura")
        self.signature_controls_container = QWidget()
        form_sign = QFormLayout(self.signature_controls_container)
        
        self.signature_file_button = QPushButton("Carregar Imagem...")
        self.signature_path_label = QLabel("Nenhuma imagem selecionada")
        self.signature_path_label.setStyleSheet("font-size: 10px; color: #aaa;")
        self.signature_size_input = QSpinBox(); self.signature_size_input.setRange(50, 1000); self.signature_size_input.setValue(300); self.signature_size_input.setSuffix(" px")
        self.signature_pos_x_input = QSpinBox(); self.signature_pos_x_input.setRange(0, 4000); self.signature_pos_x_input.setValue(1200)
        self.signature_pos_y_input = QSpinBox(); self.signature_pos_y_input.setRange(0, 4000); self.signature_pos_y_input.setValue(700)
        
        form_sign.addRow(self.signature_file_button, self.signature_path_label)
        form_sign.addRow("Tamanho:", self.signature_size_input)
        form_sign.addRow("Pos. X:", self.signature_pos_x_input)
        form_sign.addRow("Pos. Y:", self.signature_pos_y_input)
        
        layout_sign.addWidget(self.signature_checkbox)
        layout_sign.addWidget(self.signature_controls_container)
        self.signature_controls_container.setVisible(False)
        
        # Novo grupo para as Atualizações
        grp_updates = QGroupBox("Atualizações")
        layout_updates = QVBoxLayout(grp_updates)
        self.beta_checkbox = QCheckBox("Receber atualizações de teste (Pre-release)")
        
        # Leitura segura para evitar o bug de string do QSettings
        val = self.settings.value("beta_updates", False)
        salvo_beta = str(val).lower() in ['true', '1'] if isinstance(val, str) else bool(val)
        
        self.beta_checkbox.setChecked(salvo_beta)
        layout_updates.addWidget(self.beta_checkbox)
        
        layout_config.addWidget(grp_visual)
        layout_config.addWidget(grp_sign)
        layout_config.addWidget(grp_updates)
        layout_config.addStretch()

        self.tabs.addTab(tab_part, QIcon(), "👤 Participante")
        self.tabs.addTab(tab_evt, QIcon(), "📅 Evento")
        self.tabs.addTab(tab_config, QIcon(), "⚙️ Ajustes")
        self.tabs.setFixedWidth(380)

        preview_container = QGroupBox("Pré-visualização em Tempo Real")
        preview_layout = QVBoxLayout(preview_container)
        
        self.preview_label = PreviewLabel()
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        if not self.templates or not self.fonts: 
            self.preview_label.setText("⚠️  Erro: Verifique pastas 'Modelos' e 'Fontes'.")
            self.preview_label.setAlignment(Qt.AlignCenter)
        
        preview_layout.addWidget(self.preview_label)
        
        content_layout.addWidget(self.tabs)
        content_layout.addWidget(preview_container, 1)

        main_layout.addLayout(content_layout)

        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(10)
        
        self.template_excel_button = QPushButton("Baixar Modelo Excel")
        self.template_excel_button.setFlat(True)
        self.template_excel_button.setStyleSheet("text-decoration: underline; color: #aaa;")
        self.template_excel_button.setCursor(Qt.PointingHandCursor)

        self.batch_gen_button = QPushButton("Gerar em Lote (Excel/CSV)")
        self.batch_gen_button.setObjectName("btn_lote")
        self.batch_gen_button.setIcon(QIcon(os.path.join(self.icons_dir, "excel_icon.png")))

        self.generate_button = QPushButton("GERAR CERTIFICADO PDF")
        self.generate_button.setObjectName("btn_gerar")
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.setMinimumWidth(250)

        footer_layout.addWidget(self.template_excel_button)
        footer_layout.addStretch()
        footer_layout.addWidget(self.batch_gen_button)
        footer_layout.addWidget(self.generate_button)

        main_layout.addLayout(footer_layout)

        self.connect_signals()

    def connect_signals(self):
        widgets_to_update = [
            self.evento_input, self.horas_input, self.pessoa_input, self.doc_input,
            self.trabalho_input, self.funcao_custom_input, self.date_inicio_input, self.date_fim_input,
            self.pos_x_input, self.pos_y_input, self.font_size_input,
            self.signature_size_input, self.signature_pos_x_input, self.signature_pos_y_input
        ]
        for w in widgets_to_update:
            if isinstance(w, QLineEdit): w.textChanged.connect(self.update_preview)
            elif isinstance(w, QSpinBox): w.valueChanged.connect(self.update_preview)

        self.doc_tipo_combo.currentTextChanged.connect(self.update_preview)
        self.template_combo.currentTextChanged.connect(self.on_template_change)
        self.font_combo.currentTextChanged.connect(self.on_font_change)
        self.atividade_combo.currentTextChanged.connect(self.on_atividade_change)
        self.funcao_combo.currentTextChanged.connect(self.on_funcao_change)
        self.italic_checkbox.stateChanged.connect(self.update_preview)
        
        self.date_checkbox.stateChanged.connect(self.toggle_date_field)
        
        self.signature_checkbox.stateChanged.connect(self.toggle_signature_fields)
        self.signature_file_button.clicked.connect(self.select_signature_file)
        self.signature_size_input.valueChanged.connect(self.update_signature_from_controls)
        self.signature_pos_x_input.valueChanged.connect(self.update_signature_from_controls)
        self.signature_pos_y_input.valueChanged.connect(self.update_signature_from_controls)
        
        self.preview_label.signature_moved.connect(self.on_signature_dragged)

        self.generate_button.clicked.connect(self.handle_individual_save)
        self.template_excel_button.clicked.connect(self.generate_excel_template)
        self.batch_gen_button.clicked.connect(self.process_batch_file)
        
        self.twitter_button.clicked.connect(self.open_twitter)
        self.instagram_button.clicked.connect(self.open_instagram)
        self.bug_button.clicked.connect(self.report_bug)
        
        # Salva a preferência de atualização beta instantaneamente
        self.beta_checkbox.stateChanged.connect(
            lambda: self.settings.setValue("beta_updates", self.beta_checkbox.isChecked())
        )
    
    def on_signature_dragged(self, x, y):
        self.signature_pos_x_input.blockSignals(True)
        self.signature_pos_y_input.blockSignals(True)
        self.signature_pos_x_input.setValue(x)
        self.signature_pos_y_input.setValue(y)
        self.signature_pos_x_input.blockSignals(False)
        self.signature_pos_y_input.blockSignals(False)
        self.preview_label._update_signature_rect(x, y, self.signature_size_input.value())
        self.preview_label.update()

    def update_signature_from_controls(self):
        if self.signature_checkbox.isChecked():
            self.preview_label.update_signature(
                self.signature_path,
                self.signature_pos_x_input.value(),
                self.signature_pos_y_input.value(),
                self.signature_size_input.value()
            )

    def toggle_signature_fields(self):
        self.signature_controls_container.setVisible(self.signature_checkbox.isChecked())
        self.update_preview()

    def toggle_date_field(self):
        visible = self.date_checkbox.isChecked()
        self.date_inicio_label.setVisible(visible)
        self.date_inicio_input.setVisible(visible)
        self.date_fim_label.setVisible(visible)
        self.date_fim_input.setVisible(visible)
        if not visible: 
            self.date_inicio_input.clear()
            self.date_fim_input.clear()
        self.update_preview()
        
    def select_signature_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Assinatura", "", "Imagens (*.png *.jpg *.jpeg)")
        if file_name:
            self.signature_path = file_name
            self.signature_path_label.setText(os.path.basename(file_name))
            self.update_preview()

    def update_preview(self):
        if not self.template_path or not self.font_paths: return
        
        data = self.get_current_data(for_preview=True)
        data_para_preview = data.copy()
        data_para_preview['use_signature'] = False 

        sucesso, resultado_img_pil = gerar_imagem_certificado(**data_para_preview)

        if sucesso:
            pixmap = QPixmap.fromImage(ImageQt(resultado_img_pil))
            scaled_pixmap = pixmap.scaled(self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            self.preview_label.set_base_pixmap(scaled_pixmap, resultado_img_pil.size)
            
            if self.signature_checkbox.isChecked():
                self.update_signature_from_controls()
            else:
                self.preview_label.clear_signature()
        else:
            self.preview_label.setText(str(resultado_img_pil))
            self.preview_label.setPixmap(QPixmap())

    def get_documents_path(self):
        if sys.platform == "win32":
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)
            if buf.value: return buf.value
        return os.path.join(os.path.expanduser('~'), 'Documents')
        
    def setup_user_directories(self):
        try:
            docs_path = self.get_documents_path()
            app_data_path = os.path.join(docs_path, 'Gerador de Certificados')
            self.models_dir = os.path.join(app_data_path, 'Modelos')
            self.fonts_dir = os.path.join(app_data_path, 'Fontes')
            self.icons_dir = os.path.join(app_data_path, 'Icones')
            
            os.makedirs(self.models_dir, exist_ok=True)
            os.makedirs(self.fonts_dir, exist_ok=True)
            os.makedirs(self.icons_dir, exist_ok=True)
            
            default_dirs = {'Modelos': self.models_dir, 'Fontes': self.fonts_dir, 'Icones': self.icons_dir}
            for dir_name, dest_path in default_dirs.items():
                source_path = get_asset_path(dir_name)
                if os.path.exists(source_path):
                    for item in os.listdir(source_path):
                        s_item = os.path.join(source_path, item)
                        d_item = os.path.join(dest_path, item)
                        if not os.path.exists(d_item):
                            if os.path.isdir(s_item): shutil.copytree(s_item, d_item)
                            else: shutil.copy2(s_item, d_item)
        except Exception as e:
            QMessageBox.critical(self, "Erro Inicialização", f"Falha ao criar pastas.\n{e}")
        
    def open_twitter(self): QDesktopServices.openUrl(QUrl("https://x.com/yanndezedias/"))
    def open_instagram(self): QDesktopServices.openUrl(QUrl("http://instagram.com/yanndezedias/"))
    def report_bug(self):
        if QMessageBox.question(self, 'Reportar Bug', "Reportar via email?", QMessageBox.Yes|QMessageBox.No) == QMessageBox.Yes:
            QDesktopServices.openUrl(QUrl("mailto:yanndezedias16@gmail.com?subject=Bug Report - Gerador"))

    def load_resources(self, directory, formats):
        if not os.path.isdir(directory): return []
        return [f for f in os.listdir(directory) if f.lower().endswith(formats)] or []
        
    def load_fonts(self):
        fonts_dict = {}
        if not os.path.isdir(self.fonts_dir): return fonts_dict
        for f_fam in os.listdir(self.fonts_dir):
            f_path = os.path.join(self.fonts_dir, f_fam)
            if os.path.isdir(f_path):
                reg, ita = None, None
                for file in os.listdir(f_path):
                    if 'italic' in file.lower() and file.endswith('.ttf'): ita = os.path.join(f_path, file)
                    elif 'regular' in file.lower() and file.endswith('.ttf'): reg = os.path.join(f_path, file)
                if reg: fonts_dict[f_fam] = {'regular': reg, 'italic': ita if ita else reg}
        return fonts_dict
        
    def on_template_change(self, template_name): 
        if template_name: self.template_path = os.path.join(self.models_dir, template_name); self.update_preview()
    def on_font_change(self, font_name): 
        if font_name in self.fonts: self.font_paths = self.fonts[font_name]; self.update_preview()
        
    def on_atividade_change(self, atividade):
        self.evento_input.setEnabled(atividade not in self.atividades_sem_nome)
        if not self.evento_input.isEnabled(): self.evento_input.clear()
        self.update_preview()
        
    def on_funcao_change(self, funcao):
        is_apr = (funcao == "Apresentador(a)")
        is_out = (funcao == "Outro")
        self.trabalho_label.setVisible(is_apr); self.trabalho_input.setVisible(is_apr)
        self.funcao_custom_input.setVisible(is_out)
        if not is_apr: self.trabalho_input.clear()
        if is_out: self.horas_input.clear(); self.horas_input.setPlaceholderText("...")
        else: self.horas_input.setText(str(self.funcoes_horas.get(funcao, "")))
        self.update_preview()
        
    def get_funcao_participante(self):
        f = self.funcao_combo.currentText()
        return self.funcao_custom_input.text() if f == "Outro" else f
        
    def get_current_data(self, for_preview=False):
        data = {
            "nome": self.pessoa_input.text() or ("[Nome da Pessoa]" if for_preview else ""),
            "funcao_participante": self.get_funcao_participante() or ("[Função]" if for_preview else ""),
            "tipo_atividade": self.atividade_combo.currentText(),
            "nome_evento": self.evento_input.text() or ("[Nome do Evento]" if for_preview else ""),
            "carga_horaria": self.horas_input.text(),
            "template_path": self.template_path,
            "doc_tipo": self.doc_tipo_combo.currentText(),
            "doc_numero": self.doc_input.text(),
            "font_path_regular": self.font_paths['regular'],
            "font_path_italic": self.font_paths['italic'],
            "font_size": self.font_size_input.value(),
            "use_italic": self.italic_checkbox.isChecked(),
            "nome_trabalho": self.trabalho_input.text(),
            "pos_x": self.pos_x_input.value(),
            "pos_y": self.pos_y_input.value(),
            "use_date": self.date_checkbox.isChecked(),
            "data_inicio": self.date_inicio_input.text(),
            "data_fim": self.date_fim_input.text(),
            "use_signature": self.signature_checkbox.isChecked(),
            "signature_path": self.signature_path,
            "signature_pos_x": self.signature_pos_x_input.value(),
            "signature_pos_y": self.signature_pos_y_input.value(),
            "signature_size": self.signature_size_input.value()
        }
        h = data['carga_horaria']
        data['carga_horaria'] = f"{h}h" if h and not h.endswith('h') else (h or "[Horas]")
        return data

    def handle_individual_save(self):
        data = self.get_current_data()
        if data['use_signature'] and not data['signature_path']:
            QMessageBox.warning(self, "Atenção", "Selecione a imagem da assinatura.")
            return
        
        req = [data['nome'], data['tipo_atividade'], data['carga_horaria'], data['funcao_participante']]
        if self.evento_input.isEnabled(): req.append(data['nome_evento'])
        if self.trabalho_input.isVisible(): req.append(data['nome_trabalho'])
        
        if not all(req):
            QMessageBox.warning(self, "Campos Vazios", "Preencha todos os campos visíveis.")
            return
        if data['doc_tipo'] != 'Nenhum' and not data['doc_numero']:
            QMessageBox.warning(self, "Atenção", "Informe o número do documento.")
            return

        sucesso, imagem = gerar_imagem_certificado(**data)
        if sucesso:
            nome_f = data['nome'].upper().replace(' ', '_')
            func_f = data['funcao_participante'].upper().replace('(A)', '')
            default_name = f"{nome_f}-{func_f}.pdf"
            
            fname, _ = QFileDialog.getSaveFileName(self, "Salvar PDF", default_name, "PDF (*.pdf)")
            if fname:
                imagem.convert("RGB").save(fname)
                QMessageBox.information(self, "Sucesso", f"Salvo em:\n{fname}")

    def generate_excel_template(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Salvar Modelo", "modelo_importacao.xlsx", "Excel (*.xlsx)")
        if not fname: return
        df = pd.DataFrame({
            "Nome da Pessoa": ["João Silva"], "Tipo de Documento": ["CPF"], "Nº do Documento": ["000.000.000-00"],
            "Função do Participante": ["Apresentador(a)"], "Função Customizada (se Outro)": [""],
            "Nome do Trabalho (se Apresentador)": ["Impacto da IA na Educação"]
        })
        try:
            df.to_excel(fname, index=False)
            QMessageBox.information(self, "Sucesso", "Modelo criado.")
        except Exception as e: QMessageBox.critical(self, "Erro", str(e))
        
    def process_batch_file(self):
        common = self.get_current_data()
        if common['use_signature'] and not common['signature_path']:
            QMessageBox.warning(self, "Atenção", "Configure a assinatura na aba Ajustes antes.")
            return
        if not common['tipo_atividade'] or (self.evento_input.isEnabled() and not common['nome_evento']):
             QMessageBox.warning(self, "Atenção", "Preencha os dados do Evento primeiro.")
             return
             
        fname, _ = QFileDialog.getOpenFileName(self, "Abrir Planilha", "", "Planilhas (*.xlsx *.csv)")
        if not fname: return
        out_folder = QFileDialog.getExistingDirectory(self, "Pasta de Saída")
        if not out_folder: return

        try:
            df = pd.read_excel(fname) if fname.endswith('.xlsx') else pd.read_csv(fname, delimiter=';')
            df.dropna(how='all', inplace=True)
            total = len(df)
            prog = QProgressDialog("Gerando Lote...", "Cancelar", 0, total, self)
            prog.setWindowModality(Qt.WindowModal)

            for i, row in df.iterrows():
                prog.setValue(i)
                if prog.wasCanceled(): break
                
                dtipo = row.get("Tipo de Documento", "Nenhum")
                if pd.isna(dtipo) or str(dtipo).strip()=='': dtipo="Nenhum"
                
                func = row.get("Função do Participante", "")
                if func == "Outro": func = row.get("Função Customizada (se Outro)", "")
                
                ntrab = row.get("Nome do Trabalho (se Apresentador)", "")
                if pd.isna(ntrab): ntrab = ""
                
                p_data = {**common, "nome": row.get("Nome da Pessoa"), "doc_tipo": dtipo, 
                          "doc_numero": str(row.get("Nº do Documento", "")),
                          "funcao_participante": func, "nome_trabalho": ntrab}
                          
                nome_f = str(p_data['nome']).upper().replace(' ', '_')
                func_f = str(p_data['funcao_participante']).upper().replace('(A)', '')
                base = f"{nome_f}-{func_f}"
                out_name = f"{base}.pdf"
                out_p = os.path.join(out_folder, out_name)
                
                c=1
                while os.path.exists(out_p):
                    out_name = f"{base}_{c}.pdf"; out_p = os.path.join(out_folder, out_name); c+=1
                
                suc, img = gerar_imagem_certificado(**p_data)
                if suc: img.convert("RGB").save(out_p)
            
            prog.setValue(total)
            QMessageBox.information(self, "Concluído", f"{total} processados.")
        except Exception as e:
            QMessageBox.critical(self, "Erro no Lote", f"Detalhes:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    try:
        apply_stylesheet(app, theme='dark_teal.xml')
    except Exception as e:
        pass
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())