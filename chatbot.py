import os
import openai
from typing import List, Dict, Any
from pathlib import Path
import hashlib
import json
from document_processor import DocumentProcessor
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

class DocumentChatbot:
    """Chatbot especializado em responder perguntas sobre documentos"""
    
    def __init__(self):
        self.client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        self.processor = DocumentProcessor()
        self.documents = []  # Lista de documentos processados
        self.chunks = []     # Lista de chunks de texto
        self.embeddings = [] # Embeddings dos chunks
        self.vectorizer = TfidfVectorizer(max_features=1000, stop_words='english')
        self.tfidf_matrix = None
        self.conversation_history = []
        
    def process_folder(self, folder_path: str):
        """
        Processa todos os documentos em uma pasta
        
        Args:
            folder_path: Caminho para a pasta com documentos
        """
        
        folder_path = Path(folder_path)
        processed_docs = []
        all_chunks = []
        
        # Processar cada arquivo suportado
        for file_path in folder_path.rglob('*'):
            if file_path.is_file() and file_path.suffix.lower() in ['.pdf', '.docx', '.pptx']:
                try:
                    # Extrair texto
                    text_content = self.processor.extract_text(str(file_path))
                    
                    # Criar chunks
                    chunks = self.processor.create_chunks(text_content)
                    
                    # Armazenar informações do documento
                    doc_info = {
                        'filename': file_path.name,
                        'path': str(file_path),
                        'type': text_content['metadata']['document_type'],
                        'chunk_count': len(chunks),
                        'word_count': len(text_content['full_text'].split())
                    }
                    
                    processed_docs.append(doc_info)
                    
                    # Adicionar chunks com metadados
                    for i, chunk in enumerate(chunks):
                        chunk_info = {
                            'text': chunk,
                            'document': file_path.name,
                            'chunk_id': i,
                            'doc_type': text_content['metadata']['document_type']
                        }
                        all_chunks.append(chunk_info)
                    
                except Exception as e:
                    print(f"Erro ao processar {file_path.name}: {str(e)}")
                    continue
        
        # Armazenar documentos e chunks
        self.documents = processed_docs
        self.chunks = all_chunks
        
        # Criar índice de busca
        self._create_search_index()
        
        return len(processed_docs)
    
    def _create_search_index(self):
        """Cria índice de busca usando TF-IDF"""
        
        if not self.chunks:
            return
        
        # Extrair texto dos chunks
        chunk_texts = [chunk['text'] for chunk in self.chunks]
        
        # Criar matriz TF-IDF
        try:
            self.tfidf_matrix = self.vectorizer.fit_transform(chunk_texts)
        except Exception as e:
            print(f"Erro ao criar índice de busca: {str(e)}")
            self.tfidf_matrix = None
    
    def get_response(self, user_question: str) -> str:
        """
        Gera resposta para pergunta do usuário
        
        Args:
            user_question: Pergunta do usuário
            
        Returns:
            Resposta do chatbot
        """
        
        if not self.chunks:
            return "❌ Nenhum documento foi processado ainda. Por favor, faça upload de uma pasta com documentos."
        
        # Buscar chunks relevantes
        relevant_chunks = self._search_relevant_chunks(user_question)
        
        if not relevant_chunks:
            return "❌ Não encontrei informações relevantes nos documentos para responder sua pergunta."
        
        # Gerar resposta usando OpenAI
        response = self._generate_response(user_question, relevant_chunks)
        
        # Adicionar ao histórico
        self.conversation_history.append({
            'question': user_question,
            'response': response,
            'relevant_docs': [chunk['document'] for chunk in relevant_chunks[:3]]
        })
        
        return response
    
    def _search_relevant_chunks(self, query: str, top_k: int = 5) -> List[Dict[str, Any]]:
        """
        Busca chunks mais relevantes para a pergunta
        
        Args:
            query: Pergunta do usuário
            top_k: Número de chunks a retornar
            
        Returns:
            Lista de chunks relevantes
        """
        
        if self.tfidf_matrix is None:
            # Fallback: busca por palavras-chave
            return self._keyword_search(query, top_k)
        
        try:
            # Vetorizar a pergunta
            query_vector = self.vectorizer.transform([query])
            
            # Calcular similaridade
            similarities = cosine_similarity(query_vector, self.tfidf_matrix).flatten()
            
            # Pegar índices dos mais similares
            top_indices = similarities.argsort()[-top_k:][::-1]
            
            # Retornar chunks relevantes
            relevant_chunks = []
            for idx in top_indices:
                if similarities[idx] > 0.1:  # Threshold mínimo
                    chunk = self.chunks[idx].copy()
                    chunk['similarity'] = similarities[idx]
                    relevant_chunks.append(chunk)
            
            return relevant_chunks
            
        except Exception as e:
            print(f"Erro na busca: {str(e)}")
            return self._keyword_search(query, top_k)
    
    def _keyword_search(self, query: str, top_k: int = 5) -> List[Dict[str, Any]]:
        """Busca simples por palavras-chave como fallback"""
        
        query_words = query.lower().split()
        scored_chunks = []
        
        for chunk in self.chunks:
            text_lower = chunk['text'].lower()
            score = sum(word in text_lower for word in query_words)
            
            if score > 0:
                chunk_copy = chunk.copy()
                chunk_copy['similarity'] = score / len(query_words)
                scored_chunks.append(chunk_copy)
        
        # Ordenar por score e retornar top_k
        scored_chunks.sort(key=lambda x: x['similarity'], reverse=True)
        return scored_chunks[:top_k]
    
    def _generate_response(self, question: str, relevant_chunks: List[Dict[str, Any]]) -> str:
        """
        Gera resposta usando OpenAI baseada nos chunks relevantes
        
        Args:
            question: Pergunta do usuário
            relevant_chunks: Chunks de texto relevantes
            
        Returns:
            Resposta gerada
        """
        
        # Preparar contexto
        context_parts = []
        for chunk in relevant_chunks:
            context_parts.append(f"[Documento: {chunk['document']}]\n{chunk['text']}")
        
        context = "\n\n".join(context_parts)
        
        # Preparar histórico de conversa
        conversation_context = ""
        if self.conversation_history:
            recent_history = self.conversation_history[-3:]  # Últimas 3 interações
            for item in recent_history:
                conversation_context += f"Pergunta anterior: {item['question']}\n"
                conversation_context += f"Resposta anterior: {item['response']}\n\n"
        
        # Prompt para o modelo
        prompt = f"""
        Você é um assistente especializado em responder perguntas sobre documentos.
        
        CONTEXTO DOS DOCUMENTOS:
        {context}
        
        {"HISTÓRICO DA CONVERSA:" + conversation_context if conversation_context else ""}
        
        PERGUNTA ATUAL: {question}
        
        INSTRUÇÕES:
        - Responda de forma clara e objetiva
        - Use apenas informações dos documentos fornecidos
        - Se não souber a resposta, diga que não encontrou a informação
        - Cite o nome do documento quando relevante
        - Mantenha a coerência com o histórico da conversa
        - Responda em português brasileiro
        
        RESPOSTA:
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "Você é um assistente especializado em análise de documentos. Responda sempre em português brasileiro de forma clara e objetiva."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=800
            )
            
            answer = response.choices[0].message.content
            
            # Adicionar informação sobre documentos consultados
            doc_names = list(set([chunk['document'] for chunk in relevant_chunks]))
            if len(doc_names) == 1:
                footer = f"\n\n📄 *Baseado no documento: {doc_names[0]}*"
            else:
                footer = f"\n\n📄 *Baseado nos documentos: {', '.join(doc_names)}*"
            
            return answer + footer
            
        except Exception as e:
            return f"❌ Erro ao gerar resposta: {str(e)}"
    
    def get_document_summary(self) -> str:
        """Retorna resumo dos documentos processados"""
        
        if not self.documents:
            return "Nenhum documento processado."
        
        summary_parts = []
        summary_parts.append(f"📊 **Documentos processados: {len(self.documents)}**\n")
        
        total_chunks = sum(doc['chunk_count'] for doc in self.documents)
        total_words = sum(doc['word_count'] for doc in self.documents)
        
        summary_parts.append(f"📝 Total de seções: {total_chunks}")
        summary_parts.append(f"🔤 Total de palavras: {total_words:,}")
        summary_parts.append("\n**Documentos:**")
        
        for doc in self.documents:
            summary_parts.append(f"• {doc['filename']} ({doc['type']}) - {doc['word_count']:,} palavras")
        
        return "\n".join(summary_parts)
    
    def clear_history(self):
        """Limpa o histórico de conversas"""
        self.conversation_history = []
        
    def get_conversation_stats(self) -> Dict[str, Any]:
        """Retorna estatísticas da conversa"""
        
        if not self.conversation_history:
            return {"total_questions": 0}
        
        # Documentos mais consultados
        doc_mentions = {}
        for item in self.conversation_history:
            for doc in item.get('relevant_docs', []):
                doc_mentions[doc] = doc_mentions.get(doc, 0) + 1
        
        most_used_docs = sorted(doc_mentions.items(), key=lambda x: x[1], reverse=True)[:3]
        
        return {
            "total_questions": len(self.conversation_history),
            "most_consulted_docs": most_used_docs,
            "avg_docs_per_question": sum(len(item.get('relevant_docs', [])) for item in self.conversation_history) / len(self.conversation_history)
        }
