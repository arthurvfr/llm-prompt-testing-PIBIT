import json
from google import genai
from google.genai import types
from config import GEMINI_API_KEY

API_KEY = GEMINI_API_KEY  

client = genai.Client(api_key=API_KEY)

problem_description = """
Escreva uma função chamada `calcular_media(numeros)` que recebe uma lista de números inteiros e retorna a média deles. Se a lista estiver vazia, a função deve retornar 0.
"""

student_code = """
def calcular_media(numeros):
    soma = 0
    for n in numeros:
        soma += n
    
    media = soma / len(numeros)
    return media
"""

prompt_template = f"""
Você é um professor de ciência da computação experiente, especialista em pedagogia para iniciantes em programação. 
Sua tarefa é analisar o código de um aluno para um problema específico e fornecer um feedback formativo.

Regras INEGOCIÁVEIS:
1. NUNCA forneça a solução correta ou escreva o código corrigido para o aluno.
3. Se o código estiver incorreto, forneça uma dica (hint) que faça o aluno pensar e descobrir o erro por conta própria (método socrático).

Analise o seguinte problema e a submissão do aluno:

[Descrição do Problema]
{problem_description}

[Código do Aluno]
{student_code}

Retorne a sua análise no seguinte formato JSON:
{{
  "is_correct": false,
  "linha_do_erro": <numero_da_linha_se_houver_senao_null>,
  "tipo_de_erro": "<Sintaxe, Lógica ou Formatação>",
  "feedback_formativo": "<Uma explicação curta e amigável sobre o que está acontecendo de errado, sem dar a resposta>",
  "dica_acao": "<Uma pergunta ou sugestão de próximo passo para o aluno investigar>"
}}
"""

print("Analisando o código do aluno...")
try:
    response = client.models.generate_content(
        model='gemini-3-flash-preview',
        contents=prompt_template,
        config=types.GenerateContentConfig(
            #força a API a devolver exatamente um JSON válido, sem markdown
            response_mime_type="application/json",
            temperature=0.2, #temperatura baixa para respostas mais focadas e analíticas
        ),
    )
    
    #faz o parse direto do texto de resposta
    feedback_json = json.loads(response.text)
    
    print("\n--- FEEDBACK GERADO ---")
    print(f"Correto? {feedback_json.get('is_correct')}")
    print(f"Linha do Erro: {feedback_json.get('linha_do_erro')}")
    print(f"Tipo: {feedback_json.get('tipo_de_erro')}")
    print(f"Feedback: {feedback_json.get('feedback_formativo')}")
    print(f"Dica: {feedback_json.get('dica_acao')}")

except Exception as e:
    #bloco except ajustado para não chamar 'response' caso a requisição falhe
    print(f"\n[ERRO] Falha ao comunicar com a API ou ao processar o JSON: {e}")