import openai
import os
import json
from typing import Dict, List, Any
from document_processor import DocumentProcessor
import re

class AIAnalyzer:
    """Classe para análise inteligente de documentos usando OpenAI"""
    
    def __init__(self):
        self.client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        self.processor = DocumentProcessor()
    
    def analyze_document_structure(self, text_content: Dict[str, Any], filename: str) -> Dict[str, Any]:
        """
        Análise inicial do documento para identificar tipo e sugerir análises
        
        Args:
            text_content: Conteúdo extraído do documento
            filename: Nome do arquivo
            
        Returns:
            Dict com análise estrutural e sugestões
        """
        
        # Pegar uma amostra do início do documento
        full_text = text_content['full_text']
        sample_text = full_text[:3000]  # Primeiros 3000 caracteres
        
        metadata = text_content['metadata']
        
        prompt = f"""
        Analise este documento e forneça uma análise estrutural em JSON:

        FILENAME: {filename}
        TIPO: {metadata['document_type']}
        METADATA: {json.dumps(metadata, indent=2)}
        
        AMOSTRA DO TEXTO:
        {sample_text}
        
        Responda APENAS com um JSON válido contendo:
        {{
            "document_type": "tipo específico do documento (ex: contrato, relatório, apresentação, manual, etc.)",
            "main_topic": "tópico principal do documento",
            "confidence_score": "score de 0-100 da confiança na análise",
            "structure": {{
                "estimated_pages": "número estimado de páginas",
                "main_sections": ["seção1", "seção2", "seção3"],
                "has_numerical_data": true/false,
                "has_dates": true/false,
                "has_financial_info": true/false,
                "language": "idioma do documento"
            }},
            "suggested_analyses": [
                "Resumo Executivo",
                "Análise por Tópicos",
                "Extração de Dados",
                "Análise de Cláusulas",
                "Timeline de Eventos",
                "Análise Comparativa"
            ],
            "priority_analysis": "tipo de análise mais recomendada",
            "key_entities": ["entidade1", "entidade2", "entidade3"]
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=1000
            )
            
            # Tentar parsear JSON
            analysis = json.loads(response.choices[0].message.content)
            
            # Validar e limpar análise
            return self._validate_analysis(analysis, metadata)
            
        except Exception as e:
            # Fallback para análise básica
            return self._fallback_analysis(text_content, filename)
    
    def perform_detailed_analysis(self, text_content: Dict[str, Any], analysis_type: str) -> Dict[str, Any]:
        """
        Executa análise detalhada baseada no tipo escolhido
        
        Args:
            text_content: Conteúdo do documento
            analysis_type: Tipo de análise a ser executada
            
        Returns:
            Dict com resultados da análise detalhada
        """
        
        # Criar chunks do documento
        chunks = self.processor.create_chunks(text_content)
        
        # Selecionar estratégia baseada no tipo de análise
        if analysis_type == "Resumo Executivo":
            return self._executive_summary(chunks)
        elif analysis_type == "Análise por Tópicos":
            return self._topic_analysis(chunks)
        elif analysis_type == "Extração de Dados":
            return self._data_extraction(chunks)
        elif analysis_type == "Análise de Cláusulas":
            return self._clause_analysis(chunks)
        elif analysis_type == "Timeline de Eventos":
            return self._timeline_analysis(chunks)
        else:
            return self._general_analysis(chunks, analysis_type)
    
    def _executive_summary(self, chunks: List[str]) -> Dict[str, Any]:
        """Cria resumo executivo do documento"""
        
        # Processar chunks em grupos
        summaries = []
        for i, chunk in enumerate(chunks):
            prompt = f"""
            Analise este trecho do documento e crie um resumo executivo focado nos pontos principais:

            TRECHO {i+1}:
            {chunk}

            Responda com:
            - Pontos principais (máximo 3)
            - Dados importantes (números, datas, percentuais)
            - Conclusões ou recomendações
            
            Formato JSON:
            {{
                "main_points": ["ponto1", "ponto2", "ponto3"],
                "important_data": ["dado1", "dado2"],
                "conclusions": ["conclusão1", "conclusão2"]
            }}
            """
            
            try:
                response = self.client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.2,
                    max_tokens=800
                )
                
                chunk_summary = json.loads(response.choices[0].message.content)
                summaries.append(chunk_summary)
                
            except Exception as e:
                # Fallback simples
                summaries.append({
                    "main_points": [f"Resumo do trecho {i+1}"],
                    "important_data": [],
                    "conclusions": []
                })
        
        # Consolidar resumos
        return self._consolidate_summaries(summaries, "executive")
    
    def _topic_analysis(self, chunks: List[str]) -> Dict[str, Any]:
        """Análise estruturada por tópicos"""
        
        # Primeiro, identificar tópicos principais
        all_text = " ".join(chunks)
        
        prompt = f"""
        Analise este documento e identifique os principais tópicos/temas abordados:

        DOCUMENTO:
        {all_text[:5000]}...

        Responda com JSON:
        {{
            "main_topics": ["tópico1", "tópico2", "tópico3", "tópico4", "tópico5"],
            "topic_hierarchy": {{
                "tópico1": ["subtópico1", "subtópico2"],
                "tópico2": ["subtópico1", "subtópico2"]
            }}
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=600
            )
            
            topics_structure = json.loads(response.choices[0].message.content)
            
            # Analisar cada tópico
            topic_analyses = {}
            for topic in topics_structure["main_topics"]:
                topic_analyses[topic] = self._analyze_single_topic(chunks, topic)
            
            return {
                "analysis_type": "Análise por Tópicos",
                "topics_structure": topics_structure,
                "topic_analyses": topic_analyses
            }
            
        except Exception as e:
            return {"error": f"Erro na análise por tópicos: {str(e)}"}
    
    def _data_extraction(self, chunks: List[str]) -> Dict[str, Any]:
        """Extração de dados estruturados"""
        
        extracted_data = {
            "dates": [],
            "numbers": [],
            "percentages": [],
            "currencies": [],
            "names": [],
            "locations": [],
            "organizations": []
        }
        
        for chunk in chunks:
            prompt = f"""
            Extraia dados estruturados deste trecho:

            TEXTO:
            {chunk}

            Responda com JSON:
            {{
                "dates": ["data1", "data2"],
                "numbers": ["número1", "número2"],
                "percentages": ["percentual1", "percentual2"],
                "currencies": ["valor1", "valor2"],
                "names": ["nome1", "nome2"],
                "locations": ["local1", "local2"],
                "organizations": ["org1", "org2"]
            }}
            """
            
            try:
                response = self.client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.1,
                    max_tokens=500
                )
                
                chunk_data = json.loads(response.choices[0].message.content)
                
                # Consolidar dados
                for key in extracted_data.keys():
                    if key in chunk_data:
                        extracted_data[key].extend(chunk_data[key])
                        
            except Exception as e:
                continue
        
        # Remover duplicatas
        for key in extracted_data.keys():
            extracted_data[key] = list(set(extracted_data[key]))
        
        return {
            "analysis_type": "Extração de Dados",
            "extracted_data": extracted_data
        }
    
    def _clause_analysis(self, chunks: List[str]) -> Dict[str, Any]:
        """Análise específica de cláusulas (para contratos/documentos legais)"""
        
        clauses = []
        for i, chunk in enumerate(chunks):
            prompt = f"""
            Analise este trecho procurando por cláusulas, termos importantes e obrigações:

            TRECHO:
            {chunk}

            Responda com JSON:
            {{
                "clauses_found": ["cláusula1", "cláusula2"],
                "key_terms": ["termo1", "termo2"],
                "obligations": ["obrigação1", "obrigação2"],
                "important_conditions": ["condição1", "condição2"]
            }}
            """
            
            try:
                response = self.client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.2,
                    max_tokens=600
                )
                
                clause_analysis = json.loads(response.choices[0].message.content)
                clause_analysis["section"] = i + 1
                clauses.append(clause_analysis)
                
            except Exception as e:
                continue
        
        return {
            "analysis_type": "Análise de Cláusulas",
            "clauses": clauses
        }
    
    def _timeline_analysis(self, chunks: List[str]) -> Dict[str, Any]:
        """Análise de timeline/cronologia"""
        
        events = []
        for chunk in chunks:
            prompt = f"""
            Extraia eventos com datas/cronologia deste trecho:

            TEXTO:
            {chunk}

            Responda com JSON:
            {{
                "events": [
                    {{"date": "data", "event": "descrição do evento", "importance": "alta/média/baixa"}},
                    {{"date": "data", "event": "descrição do evento", "importance": "alta/média/baixa"}}
                ]
            }}
            """
            
            try:
                response = self.client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.2,
                    max_tokens=500
                )
                
                chunk_events = json.loads(response.choices[0].message.content)
                events.extend(chunk_events.get("events", []))
                
            except Exception as e:
                continue
        
        # Ordenar eventos por data (simplificado)
        events.sort(key=lambda x: x.get("date", ""))
        
        return {
            "analysis_type": "Timeline de Eventos",
            "timeline": events
        }
    
    def _general_analysis(self, chunks: List[str], analysis_type: str) -> Dict[str, Any]:
        """Análise geral personalizada"""
        
        prompt = f"""
        Realize uma análise do tipo "{analysis_type}" neste documento:

        DOCUMENTO:
        {" ".join(chunks)[:8000]}...

        Forneça uma análise estruturada e detalhada em JSON.
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=1500
            )
            
            return {
                "analysis_type": analysis_type,
                "results": response.choices[0].message.content
            }
            
        except Exception as e:
            return {"error": f"Erro na análise: {str(e)}"}
    
    def _analyze_single_topic(self, chunks: List[str], topic: str) -> Dict[str, Any]:
        """Analisa um tópico específico"""
        
        relevant_chunks = []
        topic_lower = topic.lower()
        
        # Encontrar chunks relevantes para o tópico
        for chunk in chunks:
            if topic_lower in chunk.lower():
                relevant_chunks.append(chunk)
        
        if not relevant_chunks:
            relevant_chunks = chunks[:2]  # Pegar primeiros chunks como fallback
        
        combined_text = " ".join(relevant_chunks)[:4000]
        
        prompt = f"""
        Analise este conteúdo focando especificamente no tópico "{topic}":

        CONTEÚDO:
        {combined_text}

        Responda com JSON:
        {{
            "summary": "resumo do tópico",
            "key_points": ["ponto1", "ponto2", "ponto3"],
            "details": ["detalhe1", "detalhe2"],
            "related_data": ["dado1", "dado2"]
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=600
            )
            
            return json.loads(response.choices[0].message.content)
            
        except Exception as e:
            return {
                "summary": f"Análise do tópico {topic}",
                "key_points": [],
                "details": [],
                "related_data": []
            }
    
    def _consolidate_summaries(self, summaries: List[Dict], summary_type: str) -> Dict[str, Any]:
        """Consolida múltiplos resumos em um resultado final"""
        
        all_points = []
        all_data = []
        all_conclusions = []
        
        for summary in summaries:
            all_points.extend(summary.get("main_points", []))
            all_data.extend(summary.get("important_data", []))
            all_conclusions.extend(summary.get("conclusions", []))
        
        # Consolidar usando IA
        consolidation_prompt = f"""
        Consolide estes resultados de análise em um resumo final:

        PONTOS PRINCIPAIS:
        {all_points}

        DADOS IMPORTANTES:
        {all_data}

        CONCLUSÕES:
        {all_conclusions}

        Crie um resumo executivo consolidado em JSON:
        {{
            "executive_summary": "resumo principal em 2-3 parágrafos",
            "key_findings": ["achado1", "achado2", "achado3"],
            "important_metrics": ["métrica1", "métrica2"],
            "recommendations": ["recomendação1", "recomendação2"],
            "main_conclusions": ["conclusão1", "conclusão2"]
        }}
        """
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[{"role": "user", "content": consolidation_prompt}],
                temperature=0.2,
                max_tokens=1000
            )
            
            consolidated = json.loads(response.choices[0].message.content)
            consolidated["analysis_type"] = "Resumo Executivo"
            return consolidated
            
        except Exception as e:
            return {
                "analysis_type": "Resumo Executivo",
                "executive_summary": "Resumo não disponível devido a erro de processamento",
                "key_findings": all_points[:5],
                "important_metrics": all_data[:5],
                "recommendations": [],
                "main_conclusions": all_conclusions[:3]
            }
    
    def _validate_analysis(self, analysis: Dict[str, Any], metadata: Dict[str, Any]) -> Dict[str, Any]:
        """Valida e corrige a análise retornada"""
        
        # Garantir campos obrigatórios
        required_fields = ["document_type", "main_topic", "structure", "suggested_analyses"]
        
        for field in required_fields:
            if field not in analysis:
                analysis[field] = self._get_default_value(field, metadata)
        
        # Filtrar análises baseadas no tipo de documento
        doc_type = analysis.get("document_type", "").lower()
        all_analyses = [
            "Resumo Executivo",
            "Análise por Tópicos", 
            "Extração de Dados",
            "Análise de Cláusulas",
            "Timeline de Eventos",
            "Análise Comparativa"
        ]
        
        # Personalizar análises baseadas no tipo
        if "contrato" in doc_type or "legal" in doc_type:
            analysis["suggested_analyses"] = ["Análise de Cláusulas", "Resumo Executivo", "Extração de Dados"]
        elif "relatório" in doc_type or "report" in doc_type:
            analysis["suggested_analyses"] = ["Resumo Executivo", "Análise por Tópicos", "Extração de Dados"]
        elif "apresentação" in doc_type or "presentation" in doc_type:
            analysis["suggested_analyses"] = ["Resumo Executivo", "Análise por Tópicos", "Timeline de Eventos"]
        else:
            analysis["suggested_analyses"] = all_analyses[:4]
        
        return analysis
    
    def _get_default_value(self, field: str, metadata: Dict[str, Any]) -> Any:
        """Retorna valores padrão para campos obrigatórios"""
        
        defaults = {
            "document_type": metadata.get("document_type", "Documento"),
            "main_topic": "Tópico não identificado",
            "structure": {
                "estimated_pages": metadata.get("total_pages", "N/A"),
                "main_sections": ["Seção 1", "Seção 2"],
                "has_numerical_data": False,
                "has_dates": False,
                "has_financial_info": False,
                "language": "português"
            },
            "suggested_analyses": ["Resumo Executivo", "Análise por Tópicos"]
        }
        
        return defaults.get(field, "N/A")
    
    def _fallback_analysis(self, text_content: Dict[str, Any], filename: str) -> Dict[str, Any]:
        """Análise fallback quando a IA falha"""
        
        metadata = text_content["metadata"]
        
        return {
            "document_type": metadata["document_type"],
            "main_topic": "Análise automática",
            "confidence_score": 50,
            "structure": {
                "estimated_pages": metadata.get("total_pages", "N/A"),
                "main_sections": ["Início", "Meio", "Fim"],
                "has_numerical_data": True,
                "has_dates": True,
                "has_financial_info": False,
                "language": "português"
            },
            "suggested_analyses": [
                "Resumo Executivo",
                "Análise por Tópicos",
                "Extração de Dados"
            ],
            "priority_analysis": "Resumo Executivo",
            "key_entities": []
        }
