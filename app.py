import streamlit as st
import os
from pathlib import Path
import tempfile
import zipfile
from document_processor import DocumentProcessor
from ai_analyzer import AIAnalyzer
from excel_generator import ExcelGenerator
from chatbot import DocumentChatbot
import pandas as pd

# Configuração da página
st.set_page_config(
    page_title="Analisador Inteligente de Documentos",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
st.markdown("""
<style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .feature-card {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 4px solid #667eea;
        margin: 1rem 0;
    }
    .analysis-result {
        background: #e8f4f8;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #bee5eb;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📄 Analisador Inteligente de Documentos</h1>
        <p>Processamento automático de PDFs, Word e PowerPoint com IA</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar - Configurações
    with st.sidebar:
        st.header("⚙️ Configurações")
        
        # Campo para API Key
        api_key = st.text_input(
            "🔑 OpenAI API Key",
            type="password",
            help="Insira sua chave API da OpenAI"
        )
        
        if api_key:
            os.environ["OPENAI_API_KEY"] = api_key
            st.success("✅ API Key configurada!")
        else:
            st.warning("⚠️ Configure sua API Key para continuar")
        
        st.divider()
        
        # Seleção do modo
        mode = st.selectbox(
            "🎯 Modo de Operação",
            ["📄 Análise de Documento", "🤖 Chatbot de Pasta"]
        )
    
    # Conteúdo principal
    if not api_key:
        st.info("👈 Configure sua API Key na barra lateral para começar.")
        return
    
    if mode == "📄 Análise de Documento":
        document_analysis_mode()
    else:
        chatbot_mode()

def document_analysis_mode():
    """Modo de análise de documentos"""
    st.header("📄 Análise Inteligente de Documentos")
    
    # Upload de arquivo
    uploaded_file = st.file_uploader(
        "📁 Selecione um documento",
        type=['pdf', 'docx', 'pptx'],
        help="Suporte para PDFs, Word (.docx) e PowerPoint (.pptx)"
    )
    
    if uploaded_file is not None:
        # Informações do arquivo
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📄 Arquivo", uploaded_file.name)
        with col2:
            st.metric("📊 Tamanho", f"{uploaded_file.size / 1024:.1f} KB")
        with col3:
            file_type = uploaded_file.name.split('.')[-1].upper()
            st.metric("📝 Tipo", file_type)
        
        # Resetar estado se arquivo mudou
        if 'current_file' not in st.session_state or st.session_state.current_file != uploaded_file.name:
            st.session_state.current_file = uploaded_file.name
            st.session_state.processing_state = 'initial'
            st.session_state.document_analysis = None
            st.session_state.text_content = None
        
        # Processo de análise
        if st.session_state.get('processing_state', 'initial') == 'initial':
            if st.button("🔍 Analisar Documento", type="primary"):
                process_document(uploaded_file)
        else:
            # Mostrar análise já processada
            process_document(uploaded_file)

def process_document(uploaded_file):
    """Processa o documento carregado"""
    
    # Usar session state para manter o estado do processamento
    if 'processing_state' not in st.session_state:
        st.session_state.processing_state = 'initial'
    
    if 'document_analysis' not in st.session_state:
        st.session_state.document_analysis = None
    
    if 'text_content' not in st.session_state:
        st.session_state.text_content = None
    
    # Etapa inicial: processar documento
    if st.session_state.processing_state == 'initial':
        with st.spinner("🔄 Processando documento..."):
            
            # Salvar arquivo temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
            
            try:
                # Inicializar processadores
                processor = DocumentProcessor()
                analyzer = AIAnalyzer()
                
                # Etapa 1: Extrair texto
                st.info("📖 Extraindo texto do documento...")
                text_content = processor.extract_text(tmp_path)
                st.session_state.text_content = text_content
                
                # Etapa 2: Análise inicial
                st.info("🧠 Analisando conteúdo com IA...")
                document_analysis = analyzer.analyze_document_structure(text_content, uploaded_file.name)
                st.session_state.document_analysis = document_analysis
                
                # Mudar estado
                st.session_state.processing_state = 'analyzed'
                
            except Exception as e:
                st.error(f"❌ Erro ao processar documento: {str(e)}")
                return
            finally:
                # Limpar arquivo temporário
                if 'tmp_path' in locals():
                    os.unlink(tmp_path)
            
            # Recarregar página para mostrar análise
            st.rerun()
    
    # Mostrar análise e opções
    if st.session_state.processing_state == 'analyzed' and st.session_state.document_analysis:
        # Mostrar análise inicial
        st.markdown("### 📋 Análise Inicial")
        display_document_analysis(st.session_state.document_analysis)
        
        # Etapa 3: Opções de análise
        st.markdown("### 🎯 Escolha o tipo de análise")
        analysis_options = st.session_state.document_analysis.get('suggested_analyses', [])
        
        if analysis_options:
            # Usar um form para evitar reprocessamento
            with st.form("analysis_form"):
                selected_analysis = st.selectbox(
                    "Selecione a análise desejada:",
                    analysis_options,
                    format_func=lambda x: f"🔍 {x}"
                )
                
                submitted = st.form_submit_button("📊 Executar Análise Detalhada", type="primary")
                
                if submitted:
                    perform_detailed_analysis(
                        st.session_state.text_content, 
                        selected_analysis, 
                        uploaded_file.name
                    )

def display_document_analysis(analysis):
    """Exibe a análise inicial do documento"""
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📄 Tipo de Documento")
        st.info(analysis.get('document_type', 'Não identificado'))
        
        st.markdown("#### 📊 Estrutura")
        structure = analysis.get('structure', {})
        st.write(f"**Páginas estimadas:** {structure.get('estimated_pages', 'N/A')}")
        st.write(f"**Seções principais:** {len(structure.get('main_sections', []))}")
    
    with col2:
        st.markdown("#### 🎯 Análises Sugeridas")
        suggestions = analysis.get('suggested_analyses', [])
        for i, suggestion in enumerate(suggestions, 1):
            st.write(f"{i}. {suggestion}")

def perform_detailed_analysis(text_content, analysis_type, analyzer, excel_gen, filename):
    """Executa análise detalhada"""
    with st.spinner(f"🔄 Executando {analysis_type}..."):
        
        # Análise detalhada
        detailed_results = analyzer.perform_detailed_analysis(text_content, analysis_type)
        
        # Gerar Excel
        excel_file = excel_gen.generate_excel(detailed_results, analysis_type, filename)
        
        # Mostrar resultados
        st.success("✅ Análise concluída!")
        
        # Download do Excel
        with open(excel_file, 'rb') as f:
            st.download_button(
                label="📥 Baixar Relatório Excel",
                data=f.read(),
                file_name=f"analise_{filename}_{analysis_type.lower().replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Limpar arquivo temporário
        os.unlink(excel_file)

def chatbot_mode():
    """Modo chatbot de pasta"""
    st.header("🤖 Chatbot de Documentos")
    
    # Upload de pasta (como ZIP)
    st.markdown("### 📁 Upload da Pasta")
    uploaded_zip = st.file_uploader(
        "Selecione um arquivo ZIP com seus documentos",
        type=['zip'],
        help="Faça upload de uma pasta compactada (.zip) contendo seus documentos"
    )
    
    if uploaded_zip is not None:
        if st.button("🔄 Processar Pasta"):
            setup_chatbot(uploaded_zip)
    
    # Interface do chatbot
    if "chatbot" in st.session_state:
        st.markdown("### 💬 Chat com seus Documentos")
        
        # Histórico de conversas
        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
        
        # Mostrar histórico
        for message in st.session_state.chat_history:
            with st.chat_message(message["role"]):
                st.write(message["content"])
        
        # Input do usuário
        user_input = st.chat_input("Digite sua pergunta sobre os documentos...")
        
        if user_input:
            # Adicionar mensagem do usuário
            st.session_state.chat_history.append({"role": "user", "content": user_input})
            
            # Gerar resposta
            with st.spinner("🤔 Pensando..."):
                response = st.session_state.chatbot.get_response(user_input)
                st.session_state.chat_history.append({"role": "assistant", "content": response})
            
            # Rerun para mostrar a nova mensagem
            st.rerun()

def setup_chatbot(uploaded_zip):
    """Configura o chatbot com os documentos da pasta"""
    with st.spinner("🔄 Processando documentos da pasta..."):
        
        # Extrair ZIP
        with tempfile.TemporaryDirectory() as temp_dir:
            with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Inicializar chatbot
            chatbot = DocumentChatbot()
            chatbot.process_folder(temp_dir)
            
            # Salvar no session state
            st.session_state.chatbot = chatbot
            
            st.success("✅ Chatbot configurado! Agora você pode fazer perguntas sobre seus documentos.")

if __name__ == "__main__":
    main()
