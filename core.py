from PIL import Image, ImageDraw, ImageFont

def draw_mixed_style_text(draw, pos, text_segments, fonts, max_width):
    """
    Desenha um bloco de texto centralizado com múltiplos estilos (regular, itálico)
    e com quebra de linha automática.
    """
    space_width = draw.textlength(' ', font=fonts['regular'])
    
    words_with_style = []
    for text, style in text_segments:
        font = fonts.get(style, fonts['regular']) # Usa regular se o estilo não for encontrado
        for word in text.split():
            words_with_style.append({'text': word, 'font': font, 'width': draw.textlength(word, font=font)})

    lines = []
    current_line = []
    current_width = 0
    for word_info in words_with_style:
        if not current_line or (current_width + space_width + word_info['width']) <= max_width:
            current_line.append(word_info)
            current_width += word_info['width'] + (space_width if current_line else 0)
        else:
            lines.append(current_line)
            current_line = [word_info]
            current_width = word_info['width']
    if current_line:
        lines.append(current_line)

    y = pos[1]
    line_height = fonts['regular'].getbbox('A')[3] * 1.5
    
    for line in lines:
        line_width = sum(word['width'] for word in line) + (len(line) - 1) * space_width
        x = (draw.im.size[0] - line_width) / 2
        
        for word_info in line:
            draw.text((x, y), word_info['text'], font=word_info['font'], fill="black")
            x += word_info['width'] + space_width
        y += line_height

def gerar_certificado(nome, funcao_participante, tipo_atividade, nome_evento, carga_horaria, template_path, doc_tipo, doc_numero, font_path_regular, font_path_italic, font_size, use_italic):
    """
    Gera um objeto de imagem de certificado, aplicando itálico opcionalmente.
    """
    try:
        template = Image.open(template_path)
        draw = ImageDraw.Draw(template)

        nome = (nome if nome else "[Nome da Pessoa]").upper()
        funcao_participante = funcao_participante.lower() if funcao_participante else "[Função]"
        tipo_atividade = tipo_atividade if tipo_atividade else "[Tipo de Atividade]"
        carga_horaria = carga_horaria if carga_horaria else "[Horas]"

        texto_documento = ""
        if doc_tipo == "CPF" and doc_numero:
            texto_documento = f", portador(a) do CPF {doc_numero},"
        elif doc_tipo == "Matrícula" and doc_numero:
            texto_documento = f", portador(a) da matrícula nº {doc_numero},"
        else:
            texto_documento = ","

        # --- LÓGICA DE TEXTO COM ITÁLICO OPCIONAL ---
        texto_inicial = f"Certificamos que {nome}{texto_documento} participou como {funcao_participante} da atividade de {tipo_atividade.lower()} "
        texto_evento = f'"{nome_evento}"' if nome_evento else ""
        texto_final = f" com carga horária de {carga_horaria}."
        
        # Define o estilo do evento com base na checkbox
        estilo_evento = 'italic' if use_italic and nome_evento else 'regular'
        
        text_segments = [
            (texto_inicial, 'regular'),
            (texto_evento, estilo_evento),
            (texto_final, 'regular')
        ]

        fonts = {
            'regular': ImageFont.truetype(font_path_regular, size=font_size),
            'italic': ImageFont.truetype(font_path_italic, size=font_size)
        }
        
        posicao_bloco = (250, 600)
        largura_maxima = 1500
        
        draw_mixed_style_text(draw, posicao_bloco, text_segments, fonts, largura_maxima)
        
        return True, template

    except Exception as e:
        msg_erro = f"Ocorreu um erro no core:\n{e}"
        return False, msg_erro