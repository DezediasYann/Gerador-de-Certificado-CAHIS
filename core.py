import os
from PIL import Image, ImageDraw, ImageFont

def draw_mixed_style_text(draw, pos, text_segments, fonts, max_width):
    """
    Desenha um bloco de texto centralizado com múltiplos estilos, respeitando uma
    largura máxima e uma posição inicial (pos).
    """
    space_width = draw.textlength(' ', font=fonts['regular'])
    
    words_with_style = []
    for text, style in text_segments:
        if not text: continue
        font = fonts.get(style, fonts['regular'])
        for word in text.split():
            words_with_style.append({'text': word, 'font': font, 'width': draw.textlength(word, font=font)})

    lines, current_line, current_width = [], [], 0
    for word_info in words_with_style:
        if not current_line or (current_width + space_width + word_info['width']) <= max_width:
            current_line.append(word_info)
            current_width += word_info['width'] + (space_width if current_line else 0)
        else:
            lines.append(current_line); current_line = [word_info]; current_width = word_info['width']
    if current_line: lines.append(current_line)

    y = pos[1]
    line_height = fonts['regular'].getbbox('A')[3] * 1.5
    
    for line in lines:
        line_width = sum(word['width'] for word in line) + (len(line) - 1) * space_width
        x = pos[0] + (max_width - line_width) / 2
        
        for word_info in line:
            draw.text((x, y), word_info['text'], font=word_info['font'], fill="black")
            x += word_info['width'] + space_width
        y += line_height

def gerar_certificado(nome, funcao_participante, tipo_atividade, nome_evento, carga_horaria, 
                      template_path, doc_tipo, doc_numero, font_path_regular, font_path_italic, 
                      font_size, use_italic, nome_trabalho, pos_x, pos_y,
                      use_date=False, event_date="",
                      use_signature=False, signature_path=None, signature_pos_x=0, 
                      signature_pos_y=0, signature_size=300):
    """
    Gera a imagem do certificado, com lógica de texto especial para Organizadores.
    """
    try:
        template = Image.open(template_path)
        draw = ImageDraw.Draw(template)

        nome = (nome if nome else "[Nome da Pessoa]").upper()
        funcao_participante_lower = funcao_participante.lower() if funcao_participante else "[Função]"
        tipo_atividade_str = tipo_atividade if tipo_atividade else "[Tipo de Atividade]"
        carga_horaria = carga_horaria if carga_horaria else "[Horas]"
        nome_trabalho = nome_trabalho if nome_trabalho else "[Nome do Trabalho]"

        texto_documento = ""
        if doc_tipo == "CPF" and doc_numero:
            cpf_digits = ''.join(filter(str.isdigit, str(doc_numero)))
            if len(cpf_digits) == 11:
                doc_numero_formatado = f"{cpf_digits[:3]}.{cpf_digits[3:6]}.{cpf_digits[6:9]}-{cpf_digits[9:]}"
            else:
                doc_numero_formatado = doc_numero
            texto_documento = f", portador(a) do CPF {doc_numero_formatado},"
        elif doc_tipo == "Matrícula" and doc_numero:
            texto_documento = f", portador(a) da matrícula nº {doc_numero},"
        else:
            texto_documento = ""

        texto_data = f", ocorrida em {event_date}," if use_date and event_date.strip() else ""

        if funcao_participante_lower == "apresentador(a)" and nome_trabalho != "[Nome do Trabalho]":
            texto_inicial = f"Certificamos que {nome}{texto_documento} apresentou o trabalho "
            texto_trabalho = f'"{nome_trabalho}"'
            nome_evento_formatado = f'"{nome_evento}"' if nome_evento else ""
            
            partes_finais = [f"em nosso(a) {tipo_atividade_str.lower()}", nome_evento_formatado]
            texto_final_base = " ".join(filter(None, partes_finais))
            texto_final = f"{texto_final_base}{texto_data} com carga horária de {carga_horaria}."

            estilo_trabalho = 'italic' if use_italic else 'regular'
            text_segments = [(texto_inicial, 'regular'), (texto_trabalho, estilo_trabalho), (texto_final, 'regular')]

        elif funcao_participante_lower == "organizador(a)":
            texto_documento = texto_documento or ","
            texto_inicial = f"Certificamos que {nome}{texto_documento} participou da atividade de {tipo_atividade_str.lower()} "
            texto_evento = f'"{nome_evento}"' if nome_evento else ""
            texto_final = f" na qualidade de comissão organizadora, com carga horária de {carga_horaria}."
            estilo_evento = 'italic' if use_italic and nome_evento else 'regular'
            text_segments = [(texto_inicial, 'regular'), (texto_evento, estilo_evento), (texto_data, 'regular'), (texto_final, 'regular')]

        else:
            texto_documento = texto_documento or ","
            texto_inicial = f"Certificamos que {nome}{texto_documento} participou como {funcao_participante_lower} da atividade de {tipo_atividade_str.lower()} "
            texto_evento = f'"{nome_evento}"' if nome_evento else ""
            texto_final = f" com carga horária de {carga_horaria}."
            estilo_evento = 'italic' if use_italic and nome_evento else 'regular'
            text_segments = [(texto_inicial, 'regular'), (texto_evento, estilo_evento), (texto_data, 'regular'), (texto_final, 'regular')]

        fonts = {'regular': ImageFont.truetype(font_path_regular, size=font_size), 'italic': ImageFont.truetype(font_path_italic, size=font_size)}
        
        posicao_bloco = (pos_x, pos_y)
        largura_maxima = 1500
        
        draw_mixed_style_text(draw, posicao_bloco, text_segments, fonts, largura_maxima)
        
        if use_signature and signature_path and os.path.exists(signature_path):
            try:
                signature_img = Image.open(signature_path)
                ratio = signature_img.height / signature_img.width
                new_width = signature_size
                new_height = int(new_width * ratio)
                signature_img = signature_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                mask = None
                if 'A' in signature_img.getbands():
                    mask = signature_img.split()[3]

                template.paste(signature_img, (signature_pos_x, signature_pos_y), mask)
            except Exception as e:
                print(f"Erro ao processar assinatura: {e}")
        
        return True, template

    except Exception as e:
        msg_erro = f"Ocorreu um erro no core:\n{e}"
        return False, msg_erro