import streamlit as st
import fitz  # PyMuPDF
from PIL import Image, ImageDraw
import io
import numpy as np

def pdf_to_image(pdf_file, page_number):
    """Converte uma página específica do PDF para imagem"""
    try:
        # Abre o PDF
        pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
        
        # Verifica se o número da página é válido
        if page_number < 1 or page_number > len(pdf_document):
            st.error(f"Número da página inválido. O PDF tem {len(pdf_document)} páginas.")
            return None, None
        
        # Pega a página (índice começa em 0)
        page = pdf_document[page_number - 1]
        
        # Converte para imagem
        mat = fitz.Matrix(2.0, 2.0)  # Zoom de 2x para melhor qualidade
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        # Converte para PIL Image
        image = Image.open(io.BytesIO(img_data))
        
        pdf_document.close()
        return image, (image.width, image.height)
        
    except Exception as e:
        st.error(f"Erro ao processar o PDF: {str(e)}")
        return None, None

def draw_crop_preview(image, x1, y1, x2, y2):
    """Desenha uma prévia da área de recorte na imagem"""
    # Cria uma cópia da imagem
    preview_image = image.copy()
    draw = ImageDraw.Draw(preview_image)
    
    # Desenha um retângulo vermelho na área de recorte
    draw.rectangle([x1, y1, x2, y2], outline="red", width=3)
    
    return preview_image

def crop_image(image, x1, y1, x2, y2):
    """Recorta a imagem nas coordenadas especificadas"""
    try:
        # Garante que as coordenadas estão dentro dos limites da imagem
        x1 = max(0, min(x1, image.width))
        y1 = max(0, min(y1, image.height))
        x2 = max(0, min(x2, image.width))
        y2 = max(0, min(y2, image.height))
        
        # Garante que x2 > x1 e y2 > y1
        if x2 <= x1 or y2 <= y1:
            st.error("Coordenadas inválidas. Certifique-se de que x2 > x1 e y2 > y1")
            return None
        
        # Recorta a imagem
        cropped = image.crop((x1, y1, x2, y2))
        return cropped
        
    except Exception as e:
        st.error(f"Erro ao recortar a imagem: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="PDF Crop Tool",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 PDF Crop Tool")
    st.markdown("Faça upload de um PDF e teste diferentes recortes de uma página específica")
    
    # Upload do arquivo PDF
    uploaded_file = st.file_uploader(
        "Escolha um arquivo PDF", 
        type=['pdf'],
        help="Faça upload de um arquivo PDF para começar"
    )
    
    if uploaded_file is not None:
        # Input para o número da página
        page_number = st.number_input(
            "Número da página", 
            min_value=1, 
            value=1, 
            step=1,
            help="Digite o número da página que deseja recortar"
        )
        
        # Botão para processar
        if st.button("Carregar Página"):
            with st.spinner("Processando PDF..."):
                image, dimensions = pdf_to_image(uploaded_file, page_number)
                
                if image is not None:
                    # Armazena a imagem e dimensões no session state
                    st.session_state['image'] = image
                    st.session_state['dimensions'] = dimensions
                    st.success(f"Página {page_number} carregada com sucesso!")
                    st.info(f"Dimensões da página: {dimensions[0]} x {dimensions[1]} pixels")
    
    # Se temos uma imagem carregada, mostra as opções de recorte
    if 'image' in st.session_state:
        image = st.session_state['image']
        width, height = st.session_state['dimensions']
        
        st.markdown("---")
        st.subheader("🎯 Definir Área de Recorte")
        
        # Cria duas colunas para as coordenadas
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Ponto Superior Esquerdo (x1, y1):**")
            x1 = st.number_input("x1", min_value=0, max_value=width-1, value=0, key="x1")
            y1 = st.number_input("y1", min_value=0, max_value=height-1, value=0, key="y1")
            
        with col2:
            st.markdown("**Ponto Inferior Direito (x2, y2):**")
            x2 = st.number_input("x2", min_value=1, max_value=width, value=min(300, width), key="x2")
            y2 = st.number_input("y2", min_value=1, max_value=height, value=min(300, height), key="y2")
        
        # Validação das coordenadas
        if x2 <= x1 or y2 <= y1:
            st.error("⚠️ Coordenadas inválidas! Certifique-se de que x2 > x1 e y2 > y1")
        else:
            # Mostra informações sobre o recorte
            crop_width = x2 - x1
            crop_height = y2 - y1
            st.info(f"Tamanho do recorte: {crop_width} x {crop_height} pixels")
            
            # Cria duas colunas para mostrar as imagens
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**📄 Página Original com Área de Recorte:**")
                preview_image = draw_crop_preview(image, x1, y1, x2, y2)
                st.image(preview_image, use_column_width=True)
                
            with col2:
                st.markdown("**✂️ Imagem Recortada:**")
                cropped_image = crop_image(image, x1, y1, x2, y2)
                if cropped_image is not None:
                    st.image(cropped_image, use_column_width=True)
                    
                    # Informações sobre o recorte
                    st.markdown(f"""
                    **Informações do Recorte:**
                    - Coordenadas: ({x1}, {y1}) → ({x2}, {y2})
                    - Tamanho: {crop_width} x {crop_height} pixels
                    - Área: {crop_width * crop_height:,} pixels²
                    """)
    
    # Informações sobre como usar
    with st.expander("ℹ️ Como usar"):
        st.markdown("""
        1. **Faça upload** de um arquivo PDF
        2. **Digite o número da página** que deseja recortar
        3. **Clique em "Carregar Página"** para processar
        4. **Ajuste as coordenadas** para definir a área de recorte:
           - **(x1, y1)**: Ponto superior esquerdo
           - **(x2, y2)**: Ponto inferior direito
        5. **Visualize** o resultado em tempo real
        
        **Dica:** As coordenadas são em pixels. O ponto (0,0) fica no canto superior esquerdo da imagem.
        """)

if __name__ == "__main__":
    main()
