import sys
import os
import shutil
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, 
    QPushButton, QVBoxLayout, QHBoxLayout, QFormLayout, 
    QFileDialog, QComboBox, QMessageBox, QProgressDialog, QSpinBox, QCheckBox
)
from PySide6.QtGui import QPixmap, QIcon, QDesktopServices
from PySide6.QtCore import Qt, QUrl, QSize
from PIL.ImageQt import ImageQt

from core import gerar_certificado as gerar_imagem_certificado

def get_asset_path(relative_path):
    """
    Obtém o caminho absoluto para um recurso, funcionando tanto em modo de
    desenvolvimento quanto em um aplicativo compilado pelo PyInstaller.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        # Em modo de desenvolvimento, o caminho base é a pasta do script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        # --- Configuração Inicial de Pastas do Usuário ---
        self.setup_user_directories()

        self.setWindowTitle("Gerador de Certificados")
        self.setMinimumSize(1024, 768)
        
        # (Dicionários e listas permanecem os mesmos)
        self.funcoes_horas = { 'Ouvinte': 5, 'Palestrante': 2, 'Apresentador(a)': 10, 'Organizador(a)': 10, 'Mediador(a)': 2, 'Debatedor(a)': 2, 'Outro': '' }
        self.tipos_de_atividade = [ "Palestra", "Mesa redonda", "Apresentação de trabalho", "Curso", "Oficina", "Projeto de extensão", "Evento científico", "Disciplina não curricular", "Atividade Institucionalizada", "Estágio extracurricular", "Curso de língua estrangeira", "Concurso de monografia", "Bolsa de Iniciação Científica", "Competição esportiva", "Outro" ]
        self.atividades_sem_nome = ["Disciplina não curricular", "Bolsa de Iniciação Científica"]

        # Carrega os recursos a partir da nova pasta em "Documentos"
        self.templates = self.load_resources(self.models_dir, ('.png', '.jpg', '.jpeg'))
        self.template_path = os.path.join(self.models_dir, self.templates[0]) if self.templates else ""
        
        self.fonts = self.load_fonts()
        self.font_paths = list(self.fonts.values())[0] if self.fonts else {}
        
        # --- Layouts e Widgets ---
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
        self.evento_input = QLineEdit()
        self.italic_checkbox = QCheckBox("Nome do evento em itálico")
        self.atividade_combo = QComboBox(); self.atividade_combo.addItems(self.tipos_de_atividade)
        self.horas_input = QLineEdit()
        left_form.addRow("Selecionar Modelo:", self.template_combo); left_form.addRow("Fonte:", self.font_combo); left_form.addRow("Tamanho da Fonte:", self.font_size_input)
        left_form.addRow("Nome do Evento:", self.evento_input); left_form.addRow("", self.italic_checkbox); left_form.addRow("Tipo de Atividade:", self.atividade_combo); left_form.addRow("Horas:", self.horas_input)

        self.pessoa_input = QLineEdit()
        self.doc_tipo_combo = QComboBox(); self.doc_tipo_combo.addItems(['Nenhum', 'CPF', 'Matrícula']); self.doc_input = QLineEdit()
        self.funcao_combo = QComboBox(); self.funcao_combo.addItems(self.funcoes_horas.keys())
        self.funcao_custom_input = QLineEdit(); self.funcao_custom_input.setVisible(False)
        right_form.addRow("Nome da Pessoa:", self.pessoa_input); right_form.addRow("Tipo de Documento:", self.doc_tipo_combo)
        right_form.addRow("Nº do Documento:", self.doc_input); right_form.addRow("Função do Participante:", self.funcao_combo); right_form.addRow("", self.funcao_custom_input)

        self.preview_label = QLabel(); self.preview_label.setAlignment(Qt.AlignCenter); self.preview_label.setStyleSheet("background-color: #333; color: white; font-size: 16px; border-radius: 5px;"); self.preview_label.setMinimumSize(400, 250)
        
        columns_layout.addWidget(left_container, 0, Qt.AlignCenter); columns_layout.addWidget(self.preview_label, 1, Qt.AlignCenter); columns_layout.addWidget(right_container, 0, Qt.AlignCenter)
        main_layout.addLayout(header_layout); main_layout.addLayout(columns_layout, 1)
        
        footer_layout = QHBoxLayout()
        self.generate_button = QPushButton("Gerar Certificado (Individual)"); self.generate_button.setStyleSheet("background-color: #006400; color: white; font-size: 14px; padding: 10px;")
        self.template_excel_button = QPushButton("Gerar Modelo Excel"); self.template_excel_button.setStyleSheet("font-size: 14px; padding: 10px;")
        self.batch_gen_button = QPushButton("Gerar em Lote"); self.batch_gen_button.setStyleSheet("font-size: 14px; padding: 10px; background-color: #DAA520;")
        
        footer_layout.addStretch(1); footer_layout.addWidget(self.generate_button); footer_layout.addWidget(self.template_excel_button); footer_layout.addWidget(self.batch_gen_button); footer_layout.addStretch(1)
        footer_layout.addWidget(self.twitter_button); footer_layout.addWidget(self.instagram_button); footer_layout.addWidget(self.bug_button)
        main_layout.addLayout(footer_layout); main_layout.addStretch()

        if not self.templates or not self.fonts: self.preview_label.setText("Erro: Verifique se as pastas 'Modelos' e 'Fontes' existem em Documentos/Gerador de Certificados.")

        connections = {
            self.evento_input.textChanged: self.update_preview, self.horas_input.textChanged: self.update_preview,
            self.pessoa_input.textChanged: self.update_preview, self.doc_input.textChanged: self.update_preview,
            self.doc_tipo_combo.currentTextChanged: self.update_preview, self.template_combo.currentTextChanged: self.on_template_change,
            self.font_combo.currentTextChanged: self.on_font_change, self.atividade_combo.currentTextChanged: self.on_atividade_change,
            self.funcao_combo.currentTextChanged: self.on_funcao_change, self.funcao_custom_input.textChanged: self.update_preview,
            self.font_size_input.valueChanged: self.update_preview,
            self.italic_checkbox.stateChanged: self.update_preview,
            self.generate_button.clicked: self.handle_individual_save,
            self.template_excel_button.clicked: self.generate_excel_template,
            self.batch_gen_button.clicked: self.process_batch_file,
            self.twitter_button.clicked: self.open_twitter,
            self.instagram_button.clicked: self.open_instagram,
            self.bug_button.clicked: self.report_bug,
        }
        for signal, slot in connections.items(): signal.connect(slot)
        
        self.on_atividade_change(self.atividade_combo.currentText()); self.on_funcao_change(self.funcao_combo.currentText()); self.update_preview()

    def setup_user_directories(self):
        """Verifica e cria a estrutura de pastas do app na pasta Documentos do usuário."""
        try:
            # Encontra a pasta 'Documentos' do usuário de forma segura
            docs_path = os.path.join(os.path.expanduser('~'), 'Documents')
            if not os.path.exists(docs_path):
                # Fallback para o nome em português, caso 'Documents' não exista
                docs_path = os.path.join(os.path.expanduser('~'), 'Documentos')

            # Define os caminhos que o aplicativo usará
            app_data_path = os.path.join(docs_path, 'Gerador de Certificados')
            self.models_dir = os.path.join(app_data_path, 'Modelos')
            self.fonts_dir = os.path.join(app_data_path, 'Fontes')
            self.icons_dir = os.path.join(app_data_path, 'Icones')
            
            # Cria as pastas se elas não existirem
            os.makedirs(self.models_dir, exist_ok=True)
            os.makedirs(self.fonts_dir, exist_ok=True)
            os.makedirs(self.icons_dir, exist_ok=True)

            # Copia os recursos padrão (que foram empacotados com o app) para a pasta do usuário
            # A função get_asset_path encontra os arquivos tanto em modo de dev quanto compilado
            default_dirs = {'Modelos': self.models_dir, 'Fontes': self.fonts_dir, 'Icones': self.icons_dir}
            for dir_name, dest_path in default_dirs.items():
                source_path = get_asset_path(dir_name)
                if os.path.exists(source_path):
                    for item in os.listdir(source_path):
                        source_item = os.path.join(source_path, item)
                        dest_item = os.path.join(dest_path, item)
                        if not os.path.exists(dest_item): # Só copia se o arquivo/pasta não existir
                            if os.path.isdir(source_item):
                                shutil.copytree(source_item, dest_item)
                            else:
                                shutil.copy2(source_item, dest_item)
        except Exception as e:
            QMessageBox.critical(self, "Erro na Inicialização", f"Não foi possível criar as pastas de configuração em Documentos.\nErro: {e}")

    # (O restante do código, a partir daqui, permanece o mesmo das versões anteriores)
    def open_twitter(self): QDesktopServices.openUrl(QUrl("https://x.com/yanndezedias/"))
    def open_instagram(self): QDesktopServices.openUrl(QUrl("http://instagram.com/yanndezedias/"))
    def report_bug(self):
        reply = QMessageBox.question(self, 'Reportar Bug', "Você gostaria de reportar um bug?\nIsso abrirá seu cliente de email padrão.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: QDesktopServices.openUrl(QUrl("mailto:yanndezedias16@gmail.com?subject=Report de Bug - Gerador de Certificados"))
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
        if atividade in self.atividades_sem_nome: self.evento_input.setEnabled(False); self.evento_input.clear()
        else: self.evento_input.setEnabled(True)
        self.update_preview()
    def on_funcao_change(self, funcao):
        if funcao == "Outro": self.funcao_custom_input.setVisible(True); self.horas_input.clear(); self.horas_input.setPlaceholderText("Digite as horas")
        else: self.funcao_custom_input.setVisible(False); horas = self.funcoes_horas.get(funcao, ""); self.horas_input.setText(str(horas))
        self.update_preview()
    def get_funcao_participante(self):
        funcao = self.funcao_combo.currentText(); return self.funcao_custom_input.text() if funcao == "Outro" else funcao
    def get_current_data(self, for_preview=False):
        data = { "nome": self.pessoa_input.text() if not for_preview else self.pessoa_input.text() or "[Nome da Pessoa]", "funcao_participante": self.get_funcao_participante() if not for_preview else self.get_funcao_participante() or "[Função]", "tipo_atividade": self.atividade_combo.currentText(), "nome_evento": self.evento_input.text() if not for_preview else self.evento_input.text() or "[Nome do Evento]", "carga_horaria": self.horas_input.text(), "template_path": self.template_path, "doc_tipo": self.doc_tipo_combo.currentText(), "doc_numero": self.doc_input.text(), "font_path_regular": self.font_paths['regular'], "font_path_italic": self.font_paths['italic'], "font_size": self.font_size_input.value(), "use_italic": self.italic_checkbox.isChecked() }
        horas = data['carga_horaria']; data['carga_horaria'] = f"{horas}h" if horas and not horas.endswith('h') else (horas or "[Horas]"); return data
    def update_preview(self):
        if not self.template_path or not self.font_paths: return
        data = self.get_current_data(for_preview=True)
        sucesso, resultado = gerar_imagem_certificado(**data)
        if sucesso: img_qt = ImageQt(resultado); pixmap = QPixmap.fromImage(img_qt); self.preview_label.setPixmap(pixmap.scaled(self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else: self.preview_label.setText(str(resultado)); self.preview_label.setPixmap(QPixmap())
    def handle_individual_save(self):
        data = self.get_current_data()
        funcao = data['funcao_participante']; atividade = data['tipo_atividade']
        if not all([data['nome'], atividade, data['carga_horaria'], funcao]) or (self.evento_input.isEnabled() and not data['nome_evento']): QMessageBox.warning(self, "Campos Vazios", "Por favor, preencha todos os campos habilitados."); return
        if data['doc_tipo'] != 'Nenhum' and not data['doc_numero']: QMessageBox.warning(self, "Campo Obrigatório", "Por favor, preencha o Nº do Documento."); return
        sucesso, imagem = gerar_imagem_certificado(**data)
        if sucesso:
            default_filename = f"certificado_{data['nome'].replace(' ', '_')}.pdf"; file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Certificado", default_filename, "PDF Files (*.pdf)");
            if file_name: imagem.convert("RGB").save(file_name); QMessageBox.information(self, "Sucesso", f"Certificado salvo em:\n{file_name}")
    def generate_excel_template(self):
        file_name, _ = QFileDialog.getSaveFileName(self, "Salvar Modelo Excel", "modelo_participantes.xlsx", "Excel Files (*.xlsx)")
        if not file_name: return
        df = pd.DataFrame({"Nome da Pessoa": ["Fulano de Tal"], "Tipo de Documento": ["CPF"], "Nº do Documento": ["111.222.333-44"], "Função do Participante": ["Ouvinte"], "Função Customizada (se Outro)": [""]})
        try: df.to_excel(file_name, index=False, engine='openpyxl'); QMessageBox.information(self, "Sucesso", f"Modelo Excel gerado em:\n{file_name}")
        except Exception as e: QMessageBox.critical(self, "Erro", f"Não foi possível salvar o arquivo Excel.\nErro: {e}")
    def process_batch_file(self):
        common_data = self.get_current_data()
        if not all([common_data['tipo_atividade'], common_data['carga_horaria']]) or (self.evento_input.isEnabled() and not common_data['nome_evento']): QMessageBox.warning(self, "Configuração Incompleta", "Preencha os dados do evento na interface."); return
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo de Lote", "", "Planilhas (*.xlsx *.csv)")
        if not file_name: return
        output_folder = QFileDialog.getExistingDirectory(self, "Selecionar Pasta para Salvar")
        if not output_folder: return
        try:
            df = pd.read_excel(file_name) if file_name.lower().endswith('.xlsx') else pd.read_csv(file_name, delimiter=';')
            df.dropna(how='all', inplace=True)
            total = len(df); progress = QProgressDialog("Gerando certificados...", "Cancelar", 0, total, self); progress.setWindowModality(Qt.WindowModal)
            for index, row in df.iterrows():
                progress.setValue(index)
                if progress.wasCanceled(): break
                funcao = row.get("Função do Participante", ""); funcao = row.get("Função Customizada (se Outro)", "") if funcao == "Outro" else funcao
                person_data = {**common_data, "nome": row.get("Nome da Pessoa"), "doc_tipo": row.get("Tipo de Documento"), "doc_numero": str(row.get("Nº do Documento", "")), "funcao_participante": funcao}
                doc_str = f"-{person_data['doc_numero']}" if pd.notna(person_data['doc_numero']) and person_data['doc_numero'] != 'nan' else ""
                output_filename = f"{person_data['nome'].replace(' ', '_')}{doc_str}.pdf"
                output_path = os.path.join(output_folder, output_filename)
                sucesso, imagem = gerar_imagem_certificado(**person_data)
                if sucesso: imagem.convert("RGB").save(output_path)
            progress.setValue(total); QMessageBox.information(self, "Processo Concluído", f"{total} certificados foram processados.")
        except Exception as e: QMessageBox.critical(self, "Erro ao Processar Arquivo", f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # A função de setup é chamada antes da janela ser criada
    window = MainWindow()
    window.show()
    sys.exit(app.exec())