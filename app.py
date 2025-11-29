import sys
import os
import shutil
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QPushButton, QVBoxLayout, QHBoxLayout, QFormLayout,
    QFileDialog, QComboBox, QMessageBox, QProgressDialog, QSpinBox, QCheckBox
)
from PySide6.QtGui import QPixmap, QIcon, QDesktopServices, QPainter, QPen
from PySide6.QtCore import Qt, QUrl, QSize, Signal, QPoint, QRect

from PIL import Image
from PIL.ImageQt import ImageQt

from core import gerar_certificado as gerar_imagem_certificado

# Classe customizada para o preview interativo da assinatura.
class PreviewLabel(QLabel):
    # Sinal que emite a nova posição (x, y) da assinatura após ser arrastada.
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

    def set_base_pixmap(self, pixmap, original_size):
        # Define a imagem do certificado e armazena seu tamanho original.
        self.setPixmap(pixmap)
        self.full_image_size = QSize(original_size[0], original_size[1])

    def update_signature(self, signature_path, pos_x, pos_y, width):
        # Carrega a imagem da assinatura e atualiza sua posição.
        if signature_path and os.path.exists(signature_path):
            self.signature_pixmap = QPixmap(signature_path)
            self._update_signature_rect(pos_x, pos_y, width)
        else:
            self.signature_pixmap = None
        self.update()

    def clear_signature(self):
        self.signature_pixmap = None
        self.update()

    # --- Lógica de Conversão de Coordenadas ---

    def _get_displayed_pixmap_rect(self):
        # Calcula a área e posição real da imagem exibida no widget, mantendo a proporção.
        if not self.pixmap() or self.pixmap().isNull() or not self.full_image_size.isValid():
            return QRect()

        scaled_size = self.pixmap().size().scaled(self.size(), Qt.KeepAspectRatio)
        x = (self.width() - scaled_size.width()) / 2
        y = (self.height() - scaled_size.height()) / 2
        return QRect(QPoint(x, y), scaled_size)

    def _original_coords_to_widget_rect(self, ox, oy, o_width):
        # Converte coordenadas da imagem original para um retângulo no widget.
        if not self.full_image_size.isValid(): return QRect()

        displayed_rect = self._get_displayed_pixmap_rect()
        scale_ratio = displayed_rect.width() / self.full_image_size.width()
        
        sig_orig_ratio = self.signature_pixmap.height() / self.signature_pixmap.width()
        o_height = o_width * sig_orig_ratio

        wx = displayed_rect.x() + (ox * scale_ratio)
        wy = displayed_rect.y() + (oy * scale_ratio)
        w_width = o_width * scale_ratio
        w_height = o_height * scale_ratio
        
        return QRect(int(wx), int(wy), int(w_width), int(w_height))

    def _widget_pos_to_original_coords(self, w_pos):
        # Converte a posição do cursor no widget para coordenadas na imagem original.
        if not self.full_image_size.isValid(): return QPoint(0, 0)

        displayed_rect = self._get_displayed_pixmap_rect()
        if not displayed_rect.contains(w_pos):
            return None

        scale_ratio = self.full_image_size.width() / displayed_rect.width()
        relative_pos = w_pos - displayed_rect.topLeft()
        ox = relative_pos.x() * scale_ratio
        oy = relative_pos.y() * scale_ratio

        return QPoint(int(ox), int(oy))

    def _update_signature_rect(self, pos_x, pos_y, width):
        if self.signature_pixmap:
            self.signature_rect = self._original_coords_to_widget_rect(pos_x, pos_y, width)

    # --- Eventos do Mouse ---

    def mousePressEvent(self, event):
        # Inicia o arraste se o clique for sobre a assinatura.
        if event.button() == Qt.LeftButton and self.signature_pixmap and self.signature_rect.contains(event.pos()):
            self.is_dragging = True
            self.drag_offset = event.pos() - self.signature_rect.topLeft()
            self.setCursor(Qt.ClosedHandCursor)

    def mouseMoveEvent(self, event):
        # Durante o arraste, calcula a nova posição e emite o sinal 'signature_moved'.
        if self.is_dragging:
            new_top_left = event.pos() - self.drag_offset
            original_pos = self._widget_pos_to_original_coords(new_top_left)
            if original_pos:
                self.signature_moved.emit(original_pos.x(), original_pos.y())
        # Altera o cursor para indicar que a assinatura é arrastável.
        elif self.signature_pixmap and self.signature_rect.contains(event.pos()):
            if not self.is_hovering:
                self.is_hovering = True
                self.setCursor(Qt.OpenHandCursor)
        elif self.is_hovering:
            self.is_hovering = False
            self.setCursor(Qt.ArrowCursor)

    def mouseReleaseEvent(self, event):
        # Finaliza o arraste.
        if event.button() == Qt.LeftButton:
            self.is_dragging = False
            self.setCursor(Qt.OpenHandCursor if self.is_hovering else Qt.ArrowCursor)

    # --- Desenho Customizado ---

    def paintEvent(self, event):
        # Desenha primeiro o certificado (comportamento padrão do QLabel).
        super().paintEvent(event)

        # Depois, desenha a assinatura e uma caixa de marcação por cima.
        if self.signature_pixmap and not self.signature_pixmap.isNull():
            painter = QPainter(self)
            painter.drawPixmap(self.signature_rect, self.signature_pixmap)
            pen = QPen(Qt.red, 2, Qt.DashLine)
            painter.setPen(pen)
            painter.drawRect(self.signature_rect)
            painter.end()

# Janela principal da aplicação.
def get_asset_path(relative_path):
    # Encontra o caminho para assets, funcionando tanto em modo de dev quanto empacotado.
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.signature_path = None 
        
        self.setup_user_directories()
        self.setWindowTitle("Gerador de Certificados")
        self.setWindowIcon(QIcon("icone.ico"))
        self.setMinimumSize(1024, 768)
        
        self.funcoes_horas = { 'Ouvinte': 5, 'Palestrante': 2, 'Apresentador(a)': 10, 'Organizador(a)': 10, 'Mediador(a)': 2, 'Debatedor(a)': 2, 'Outro': '' }
        self.tipos_de_atividade = [ "Seminário", "Palestra", "Mesa redonda", "Apresentação de trabalho", "Curso", "Oficina", "Projeto de extensão", "Evento científico", "Disciplina não curricular", "Atividade Institucionalizada", "Estágio extracurricular", "Curso de língua estrangeira", "Concurso de monografia", "Bolsa de Iniciação Científica", "Competição esportiva", "Outro" ]
        self.atividades_sem_nome = ["Disciplina não curricular", "Bolsa de Iniciação Científica"]

        self.templates = self.load_resources(self.models_dir, ('.png', '.jpg', '.jpeg'))
        self.template_path = os.path.join(self.models_dir, self.templates[0]) if self.templates else ""
        self.fonts = self.load_fonts()
        self.font_paths = list(self.fonts.values())[0] if self.fonts else {}
        
        # --- Configuração da Interface Gráfica ---
        main_layout = QVBoxLayout(self)
        header_layout = QHBoxLayout()
        title_label = QLabel("Gerador de Certificado"); title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.twitter_button = QPushButton(); self.twitter_button.setIcon(QIcon(os.path.join(self.icons_dir, "twitter_icon.png"))); self.twitter_button.setToolTip("Abrir Twitter")
        self.instagram_button = QPushButton(); self.instagram_button.setIcon(QIcon(os.path.join(self.icons_dir, "instagram_icon.png"))); self.instagram_button.setToolTip("Abrir Instagram")
        self.bug_button = QPushButton(); self.bug_button.setIcon(QIcon(os.path.join(self.icons_dir, "bug_icon.png"))); self.bug_button.setToolTip("Reportar um bug")
        icon_buttons = [self.twitter_button, self.instagram_button, self.bug_button]
        for button in icon_buttons:
            button.setIconSize(QSize(24, 24)); button.setStyleSheet("QPushButton { border: none; background-color: transparent; }"); button.setFixedSize(QSize(32, 32))
        header_layout.addWidget(title_label); header_layout.addStretch(); header_layout.addWidget(self.twitter_button); header_layout.addWidget(self.instagram_button); header_layout.addWidget(self.bug_button)
        
        columns_layout = QHBoxLayout()
        left_container = QWidget(); left_form = QFormLayout(left_container); left_form.setContentsMargins(0,0,0,0); left_container.setMaximumWidth(350)
        right_container = QWidget(); right_form = QFormLayout(right_container); right_form.setContentsMargins(0,0,0,0); right_container.setMaximumWidth(350)

        self.template_combo = QComboBox(); self.template_combo.addItems(self.templates) if self.templates else None
        self.font_combo = QComboBox(); self.font_combo.addItems(self.fonts.keys()) if self.fonts else None
        self.font_size_input = QSpinBox(); self.font_size_input.setRange(10, 200); self.font_size_input.setValue(50); self.font_size_input.setSuffix(" pt")
        self.pos_x_input = QSpinBox(); self.pos_x_input.setRange(0, 4000); self.pos_x_input.setValue(250)
        self.pos_y_input = QSpinBox(); self.pos_y_input.setRange(0, 4000); self.pos_y_input.setValue(600)
        self.evento_input = QLineEdit()
        self.italic_checkbox = QCheckBox("Nome do evento/trabalho em itálico")
        self.atividade_combo = QComboBox(); self.atividade_combo.addItems(self.tipos_de_atividade)
        self.horas_input = QLineEdit()
        self.date_checkbox = QCheckBox("Incluir Data no Certificado")
        self.date_label = QLabel("Data do Evento (DD/MM/AAAA):")
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("DD/MM/AAAA")

        left_form.addRow("Selecionar Modelo:", self.template_combo); left_form.addRow("Fonte:", self.font_combo); left_form.addRow("Tamanho da Fonte:", self.font_size_input)
        left_form.addRow("Posição X do Texto:", self.pos_x_input); left_form.addRow("Posição Y do Texto:", self.pos_y_input)
        left_form.addRow("Tipo de Atividade:", self.atividade_combo); left_form.addRow("Nome do Evento:", self.evento_input)
        left_form.addRow("", self.italic_checkbox); left_form.addRow(self.date_checkbox); left_form.addRow(self.date_label, self.date_input)
        left_form.addRow("Horas:", self.horas_input)

        self.pessoa_input = QLineEdit()
        self.trabalho_label = QLabel("Nome do Trabalho:"); self.trabalho_input = QLineEdit()
        self.doc_tipo_combo = QComboBox(); self.doc_tipo_combo.addItems(['Nenhum', 'CPF', 'Matrícula']); self.doc_input = QLineEdit()
        self.funcao_combo = QComboBox(); self.funcao_combo.addItems(self.funcoes_horas.keys())
        self.funcao_custom_input = QLineEdit(); self.funcao_custom_input.setVisible(False)
        right_form.addRow("Nome da Pessoa:", self.pessoa_input); right_form.addRow(self.trabalho_label, self.trabalho_input)
        right_form.addRow("Tipo de Documento:", self.doc_tipo_combo); right_form.addRow("Nº do Documento:", self.doc_input); right_form.addRow("Função do Participante:", self.funcao_combo); right_form.addRow("", self.funcao_custom_input)

        self.signature_checkbox = QCheckBox("Inserir Assinatura (Experimental)")
        self.signature_controls_container = QWidget()
        signature_layout = QFormLayout(self.signature_controls_container)
        signature_layout.setContentsMargins(0, 10, 0, 0)
        self.signature_file_button = QPushButton("Selecionar Arquivo")
        self.signature_path_label = QLabel("Nenhum arquivo selecionado.")
        self.signature_path_label.setStyleSheet("font-style: italic; color: #888;")
        self.signature_size_input = QSpinBox(); self.signature_size_input.setRange(50, 1000); self.signature_size_input.setValue(300); self.signature_size_input.setSuffix(" px")
        self.signature_pos_x_input = QSpinBox(); self.signature_pos_x_input.setRange(0, 4000); self.signature_pos_x_input.setValue(1200)
        self.signature_pos_y_input = QSpinBox(); self.signature_pos_y_input.setRange(0, 4000); self.signature_pos_y_input.setValue(700)
        signature_layout.addRow(self.signature_file_button, self.signature_path_label)
        signature_layout.addRow("Largura da Assinatura:", self.signature_size_input)
        signature_layout.addRow("Posição X da Assinatura:", self.signature_pos_x_input)
        signature_layout.addRow("Posição Y da Assinatura:", self.signature_pos_y_input)
        right_form.addRow(self.signature_checkbox)
        right_form.addRow(self.signature_controls_container)
        self.signature_controls_container.setVisible(False)
        
        # Usa a classe customizada PreviewLabel para a área de preview.
        self.preview_label = PreviewLabel()
        self.preview_label.setStyleSheet("background-color: #333; color: white; font-size: 16px; border-radius: 5px;")
        self.preview_label.setMinimumSize(400, 250)
        
        columns_layout.addWidget(left_container, 0, Qt.AlignCenter); columns_layout.addWidget(self.preview_label, 1, Qt.AlignCenter); columns_layout.addWidget(right_container, 0, Qt.AlignCenter)
        main_layout.addLayout(header_layout); main_layout.addLayout(columns_layout, 1)
        
        footer_layout = QHBoxLayout()
        self.generate_button = QPushButton("Gerar Certificado (Individual)"); self.generate_button.setStyleSheet("background-color: #006400; color: white; font-size: 14px; padding: 10px;")
        self.template_excel_button = QPushButton("Gerar Modelo Excel"); self.template_excel_button.setStyleSheet("font-size: 14px; padding: 10px;")
        self.batch_gen_button = QPushButton("Gerar em Lote"); self.batch_gen_button.setStyleSheet("font-size: 14px; padding: 10px; background-color: #DAA520;")
        
        footer_layout.addStretch(1); footer_layout.addWidget(self.generate_button); footer_layout.addWidget(self.template_excel_button); footer_layout.addWidget(self.batch_gen_button); footer_layout.addStretch(1)
        main_layout.addLayout(footer_layout); main_layout.addStretch()

        if not self.templates or not self.fonts: self.preview_label.setText("Erro: Verifique pastas 'Modelos' e 'Fontes'.")
        
        # Conecta os sinais (eventos) dos widgets aos seus respectivos slots (funções).
        connections = {
            self.evento_input.textChanged: self.update_preview, self.horas_input.textChanged: self.update_preview,
            self.pessoa_input.textChanged: self.update_preview, self.doc_input.textChanged: self.update_preview,
            self.trabalho_input.textChanged: self.update_preview,
            self.pos_x_input.valueChanged: self.update_preview,
            self.pos_y_input.valueChanged: self.update_preview,
            self.doc_tipo_combo.currentTextChanged: self.update_preview, self.template_combo.currentTextChanged: self.on_template_change,
            self.font_combo.currentTextChanged: self.on_font_change, self.atividade_combo.currentTextChanged: self.on_atividade_change,
            self.funcao_combo.currentTextChanged: self.on_funcao_change, self.funcao_custom_input.textChanged: self.update_preview,
            self.font_size_input.valueChanged: self.update_preview,
            self.italic_checkbox.stateChanged: self.update_preview,
            self.date_checkbox.stateChanged: self.toggle_date_field,
            self.date_input.textChanged: self.update_preview,
            self.generate_button.clicked: self.handle_individual_save,
            self.template_excel_button.clicked: self.generate_excel_template,
            self.batch_gen_button.clicked: self.process_batch_file,
            self.twitter_button.clicked: self.open_twitter,
            self.instagram_button.clicked: self.open_instagram,
            self.bug_button.clicked: self.report_bug,
            self.signature_checkbox.stateChanged: self.toggle_signature_fields,
            self.signature_file_button.clicked: self.select_signature_file,
            self.signature_size_input.valueChanged: self.update_signature_from_controls,
            self.signature_pos_x_input.valueChanged: self.update_signature_from_controls,
            self.signature_pos_y_input.valueChanged: self.update_signature_from_controls,
        }
        for signal, slot in connections.items(): signal.connect(slot)
        
        # Conecta o sinal de arrastar da assinatura ao slot que atualiza os SpinBoxes.
        self.preview_label.signature_moved.connect(self.on_signature_dragged)
        
        self.on_atividade_change(self.atividade_combo.currentText()); self.on_funcao_change(self.funcao_combo.currentText()); self.toggle_date_field(); self.update_preview()

    def on_signature_dragged(self, x, y):
        # Atualiza os SpinBoxes quando a assinatura é arrastada.
        self.signature_pos_x_input.blockSignals(True)
        self.signature_pos_y_input.blockSignals(True)
        
        self.signature_pos_x_input.setValue(x)
        self.signature_pos_y_input.setValue(y)

        self.signature_pos_x_input.blockSignals(False)
        self.signature_pos_y_input.blockSignals(False)
        
        self.preview_label._update_signature_rect(x, y, self.signature_size_input.value())
        self.preview_label.update()

    def update_signature_from_controls(self):
        # Atualiza o preview da assinatura quando os valores dos SpinBoxes mudam.
        if self.signature_checkbox.isChecked():
            self.preview_label.update_signature(
                self.signature_path,
                self.signature_pos_x_input.value(),
                self.signature_pos_y_input.value(),
                self.signature_size_input.value()
            )

    def toggle_signature_fields(self):
        # Mostra ou esconde os controles da assinatura.
        self.signature_controls_container.setVisible(self.signature_checkbox.isChecked())
        self.update_preview()

    def toggle_date_field(self):
        visible = self.date_checkbox.isChecked()
        self.date_label.setVisible(visible)
        self.date_input.setVisible(visible)
        if not visible:
            self.date_input.clear()
        self.update_preview()
        
    def select_signature_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Assinatura", "", "Imagens (*.png *.jpg *.jpeg)")
        if file_name:
            self.signature_path = file_name
            self.signature_path_label.setText(os.path.basename(file_name))
            self.update_preview()

    def update_preview(self):
        # Função principal que redesenha o preview do certificado.
        if not self.template_path or not self.font_paths: return
        
        data = self.get_current_data(for_preview=True)

        data_para_preview = data.copy()
        data_para_preview['use_signature'] = False 

        sucesso, resultado_img_pil = gerar_imagem_certificado(**data_para_preview)

        if sucesso:
            pixmap = QPixmap.fromImage(ImageQt(resultado_img_pil))
            scaled_pixmap = pixmap.scaled(self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            # Define o fundo limpo (sem assinatura fixa)
            self.preview_label.set_base_pixmap(scaled_pixmap, resultado_img_pil.size)
            
            # Agora, se a caixa estiver marcada, o PreviewLabel desenha a assinatura flutuante por cima
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
            CSIDL_PERSONAL = 5; SHGFP_TYPE_CURRENT = 0
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
            if buf.value: return buf.value
        return os.path.join(os.path.expanduser('~'), 'Documents')
        
    def setup_user_directories(self):
        # Cria as pastas necessárias para o programa em 'Documentos'.
        try:
            docs_path = self.get_documents_path(); app_data_path = os.path.join(docs_path, 'Gerador de Certificados')
            self.models_dir, self.fonts_dir, self.icons_dir = os.path.join(app_data_path, 'Modelos'), os.path.join(app_data_path, 'Fontes'), os.path.join(app_data_path, 'Icones')
            os.makedirs(self.models_dir, exist_ok=True); os.makedirs(self.fonts_dir, exist_ok=True); os.makedirs(self.icons_dir, exist_ok=True)
            default_dirs = {'Modelos': self.models_dir, 'Fontes': self.fonts_dir, 'Icones': self.icons_dir}
            for dir_name, dest_path in default_dirs.items():
                source_path = get_asset_path(dir_name)
                if os.path.exists(source_path):
                    for item in os.listdir(source_path):
                        source_item, dest_item = os.path.join(source_path, item), os.path.join(dest_path, item)
                        if not os.path.exists(dest_item):
                            if os.path.isdir(source_item): shutil.copytree(source_item, dest_item)
                            else: shutil.copy2(source_item, dest_item)
        except Exception as e: QMessageBox.critical(self, "Erro na Inicialização", f"Não foi possível criar as pastas de configuração.\nErro: {e}")
        
    def open_twitter(self): QDesktopServices.openUrl(QUrl("https://x.com/yanndezedias/"))
    def open_instagram(self): QDesktopServices.openUrl(QUrl("http://instagram.com/yanndezedias/"))
    def report_bug(self):
        reply = QMessageBox.question(self, 'Reportar Bug', "Deseja reportar um bug?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: QDesktopServices.openUrl(QUrl("mailto:yanndezedias16@gmail.com?subject=Bug - Gerador de Certificados"))
        
    def load_resources(self, directory, formats):
        if not os.path.isdir(directory): return []
        return [f for f in os.listdir(directory) if f.lower().endswith(formats)] or []
        
    def load_fonts(self):
        fonts_dict = {}
        if not os.path.isdir(self.fonts_dir): return fonts_dict
        for font_family in os.listdir(self.fonts_dir):
            folder_path = os.path.join(self.fonts_dir, font_family)
            if os.path.isdir(folder_path):
                found_regular, found_italic = None, None
                for file in os.listdir(folder_path):
                    if 'italic' in file.lower() and file.lower().endswith('.ttf'): found_italic = os.path.join(folder_path, file)
                    elif 'regular' in file.lower() and file.lower().endswith('.ttf'): found_regular = os.path.join(folder_path, file)
                if found_regular: fonts_dict[font_family] = {'regular': found_regular, 'italic': found_italic if found_italic else found_regular}
        return fonts_dict
        
    def on_template_change(self, template_name): self.template_path = os.path.join(self.models_dir, template_name); self.update_preview()
    def on_font_change(self, font_name):
        if font_name in self.fonts: self.font_paths = self.fonts[font_name]; self.update_preview()
        
    def on_atividade_change(self, atividade):
        self.evento_input.setEnabled(atividade not in self.atividades_sem_nome)
        if not self.evento_input.isEnabled(): self.evento_input.clear()
        self.update_preview()
        
    def on_funcao_change(self, funcao):
        is_apresentador, is_outro = (funcao == "Apresentador(a)"), (funcao == "Outro")
        self.trabalho_label.setVisible(is_apresentador); self.trabalho_input.setVisible(is_apresentador)
        self.funcao_custom_input.setVisible(is_outro)
        if not is_apresentador: self.trabalho_input.clear()
        if is_outro: self.horas_input.clear(); self.horas_input.setPlaceholderText("Digite as horas")
        else: self.horas_input.setText(str(self.funcoes_horas.get(funcao, "")))
        self.update_preview()
        
    def get_funcao_participante(self):
        funcao = self.funcao_combo.currentText()
        return self.funcao_custom_input.text() if funcao == "Outro" else funcao
        
    def get_current_data(self, for_preview=False):
        # Coleta todos os dados da interface para enviar ao core.
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
            "event_date": self.date_input.text(),
            "use_signature": self.signature_checkbox.isChecked(),
            "signature_path": self.signature_path,
            "signature_pos_x": self.signature_pos_x_input.value(),
            "signature_pos_y": self.signature_pos_y_input.value(),
            "signature_size": self.signature_size_input.value()
        }
        horas = data['carga_horaria']
        data['carga_horaria'] = f"{horas}h" if horas and not horas.endswith('h') else (horas or "[Horas]")
        return data

    def handle_individual_save(self):
        # Gera e salva um único certificado.
        data = self.get_current_data()
        if data['use_signature'] and not data['signature_path']:
            QMessageBox.warning(self, "Arquivo Faltando", "Selecione uma imagem para a assinatura.")
            return
        
        campos_obrigatorios = [data['nome'], data['tipo_atividade'], data['carga_horaria'], data['funcao_participante']]
        if self.evento_input.isEnabled(): campos_obrigatorios.append(data['nome_evento'])
        if self.trabalho_input.isVisible(): campos_obrigatorios.append(data['nome_trabalho'])
        if not all(campos_obrigatorios):
            QMessageBox.warning(self, "Campos Vazios", "Preencha todos os campos habilitados.")
            return
        if data['doc_tipo'] != 'Nenhum' and not data['doc_numero']:
            QMessageBox.warning(self, "Campo Obrigatório", "Preencha o Nº do Documento.")
            return

        sucesso, imagem = gerar_imagem_certificado(**data)
        if sucesso:
            # --- ALTERAÇÃO INICIA AQUI: Nova lógica para o nome do arquivo ---
            nome_formatado = data['nome'].upper().replace(' ', '_')
            funcao_formatada = data['funcao_participante'].upper().replace('(A)', '')
            default_filename = f"{nome_formatado}-{funcao_formatada}.pdf"
            # --- FIM DA ALTERAÇÃO ---

            file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Certificado", default_filename, "PDF Files (*.pdf)")
            if file_name:
                imagem.convert("RGB").save(file_name)
                QMessageBox.information(self, "Sucesso", f"Certificado salvo em:\n{file_name}")

    def generate_excel_template(self):
        # Gera um arquivo .xlsx modelo para a geração em lote.
        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Modelo Excel", "modelo_participantes.xlsx", "Excel Files (*.xlsx)")
        if not file_name: return
        df = pd.DataFrame({
            "Nome da Pessoa": ["Fulano de Tal"], "Tipo de Documento": ["CPF"], "Nº do Documento": ["111.222.333-44"],
            "Função do Participante": ["Apresentador(a)"], "Função Customizada (se Outro)": [""],
            "Nome do Trabalho (se Apresentador)": ["Avanços na Geração Procedural"]
        })
        try:
            df.to_excel(file_name, index=False, engine='openpyxl')
            QMessageBox.information(self, "Sucesso", f"Modelo Excel gerado em:\n{file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar o arquivo Excel.\nErro: {e}")
        
    def process_batch_file(self):
        # Processa um arquivo .xlsx ou .csv para gerar múltiplos certificados.
        common_data = self.get_current_data()
        if common_data['use_signature'] and not common_data['signature_path']:
            QMessageBox.warning(self, "Arquivo Faltando", "Selecione uma imagem de assinatura antes de gerar em lote.")
            return
        if not all([common_data['tipo_atividade'], common_data['carga_horaria']]) or (self.evento_input.isEnabled() and not common_data['nome_evento']):
            QMessageBox.warning(self, "Configuração Incompleta", "Preencha os dados do evento na interface.")
            return
        
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo de Lote", "", "Planilhas (*.xlsx *.csv)")
        if not file_name: return
        output_folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta para Salvar")
        if not output_folder: return

        try:
            df_participants = pd.read_excel(file_name) if file_name.lower().endswith('.xlsx') else pd.read_csv(file_name, delimiter=';')
            df_participants.dropna(how='all', inplace=True)
            total = len(df_participants)
            progress = QProgressDialog("Gerando certificados...", "Cancelar", 0, total, self)
            progress.setWindowModality(Qt.WindowModal)

            for index, row in df_participants.iterrows():
                progress.setValue(index)
                if progress.wasCanceled(): break

                doc_tipo = row.get("Tipo de Documento", "Nenhum")
                if pd.isna(doc_tipo) or str(doc_tipo).strip() == '': doc_tipo = "Nenhum"
                funcao = row.get("Função do Participante", "")
                if funcao == "Outro": funcao = row.get("Função Customizada (se Outro)", "")
                nome_trabalho = "" if pd.isna(row.get("Nome do Trabalho (se Apresentador)")) else row.get("Nome do Trabalho (se Apresentador)", "")
                
                person_data = {**common_data,
                               "nome": row.get("Nome da Pessoa"),
                               "doc_tipo": doc_tipo,
                               "doc_numero": str(row.get("Nº do Documento", "")),
                               "funcao_participante": funcao,
                               "nome_trabalho": nome_trabalho}
                
                # --- ALTERAÇÃO INICIA AQUI: Nova lógica para o nome do arquivo em lote ---
                nome_pessoa = person_data.get("nome", "NOME_INVALIDO")
                funcao_pessoa = person_data.get("funcao_participante", "FUNCAO_INVALIDA")

                nome_formatado = nome_pessoa.upper().replace(' ', '_')
                funcao_formatada = funcao_pessoa.upper().replace('(A)', '')
                base_filename = f"{nome_formatado}-{funcao_formatada}"
                # --- FIM DA ALTERAÇÃO ---
                
                output_filename = f"{base_filename}.pdf"
                output_path = os.path.join(output_folder, output_filename)
                
                counter = 1
                while os.path.exists(output_path):
                    output_filename = f"{base_filename}_{counter}.pdf"
                    output_path = os.path.join(output_folder, output_filename)
                    counter += 1
                
                sucesso, imagem = gerar_imagem_certificado(**person_data)
                if sucesso: imagem.convert("RGB").save(output_path)
            
            progress.setValue(total)
            QMessageBox.information(self, "Processo Concluído", f"{total} certificados foram processados.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Processar Arquivo", f"Ocorreu um erro: {e}\nVerifique o formato do arquivo.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())