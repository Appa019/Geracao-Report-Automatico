# Configurações da aplicação
import os
import sys

# Adicionar diretório atual ao path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# Configurações da OpenAI
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Configurações de upload
MAX_FILE_SIZE = 200 * 1024 * 1024  # 200MB
SUPPORTED_FORMATS = ['.pdf', '.docx', '.pptx']

# Configurações de chunking
DEFAULT_CHUNK_SIZE = 4000
MAX_CHUNK_SIZE = 6000

# Configurações do modelo
GPT_MODEL_ANALYSIS = "gpt-4o-mini"
GPT_MODEL_DETAILED = "gpt-4"

# Verificar se a API key está configurada
def check_openai_config():
    """Verifica se a configuração da OpenAI está correta"""
    return OPENAI_API_KEY is not None and OPENAI_API_KEY.startswith("sk-")
