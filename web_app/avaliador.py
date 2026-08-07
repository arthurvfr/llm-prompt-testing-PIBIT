"""
Núcleo de avaliação reutilizado do script de lote (avaliar_lote.py).

Mantém o mesmo banco de problemas, o mesmo prompt Logic-of-Thought (LoT)
e a mesma chamada ao Gemini, agora expostos para uso pela aplicação web.
"""
import json
import os
import sys

from google import genai
from google.genai import types

MODEL = "gemini-flash-latest"


def _carregar_api_key():
    """Busca a chave na variável de ambiente ou no config.py da pasta raiz."""
    chave = os.environ.get("GEMINI_API_KEY")
    if chave:
        return chave

    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if raiz not in sys.path:
        sys.path.insert(0, raiz)
    try:
        from config import GEMINI_API_KEY  # noqa: PLC0415
        return GEMINI_API_KEY
    except ImportError as exc:
        raise RuntimeError(
            "GEMINI_API_KEY não encontrada. Defina a variável de ambiente "
            "GEMINI_API_KEY ou mantenha o config.py na pasta raiz do projeto."
        ) from exc


client = genai.Client(api_key=_carregar_api_key())

banco_problemas = {
    "1132": "Escreva um programa que leia dois valores inteiros X e Y. Calcule e mostre a soma de todos os números não múltiplos de 13 entre X e Y, incluindo ambos.",
    "1153": "Ler um valor N (N < 13). Calcular e escrever seu respectivo fatorial (N!). Fatorial de N = N * (N-1) * (N-2) * (N-3) * ... * 1.",
    "1828": "Bazinga! Ler número de casos (N). Ler escolhas de Sheldon e Raj. Imprimir 'Caso #i: Bazinga!' se Sheldon vence, 'Raj trapaceou!' se Raj vence, 'De novo!' se empatar. Regras: tesoura corta papel, papel cobre pedra, pedra esmaga lagarto, lagarto envenena Spock, Spock esmaga tesoura, tesoura decapita lagarto, lagarto come papel, papel refuta Spock, Spock vaporiza pedra, pedra quebra tesoura.",
    "1873": "Ler número de casos (C). Ler escolhas de rajesh e sheldon. Imprimir o nome do vencedor ('rajesh' ou 'sheldon') ou 'empate'. Segue as mesmas regras de pedra-papel-tesoura-lagarto-Spock do problema 1828.",
}

titulos_problemas = {
    "1132": "1132 - Soma de Ímpares Consecutivos I (não múltiplos de 13)",
    "1153": "1153 - Fatorial Simples",
    "1828": "1828 - Bazinga!",
    "1873": "1873 - Pedra, Papel, Tesoura, Lagarto, Spock",
}


# ---------------------------------------------------------------------------
# Casos de teste usados para verificar o código antes de gerar o feedback.
# Cada item: (entrada, saida_esperada, o_que_testa)
# As saídas foram conferidas contra soluções de referência.
# ---------------------------------------------------------------------------
casos_teste = {
    "1132": [
        ("100\n200", "13954", "caso padrão da amostra oficial"),
        ("200\n100", "13954", "X maior que Y — exige inverter os valores"),
        ("13\n13", "0", "intervalo unitário sendo múltiplo de 13"),
        ("1\n1", "1", "intervalo unitário não múltiplo"),
        ("0\n0", "0", "zero é múltiplo de 13"),
        ("1\n12", "78", "nenhum múltiplo dentro do intervalo"),
        ("12\n14", "26", "múltiplo exatamente no meio"),
        ("26\n26", "0", "múltiplo maior que 13"),
        ("-14\n-12", "-26", "números negativos"),
        ("1\n1000", "462462", "intervalo grande"),
    ],
    "1153": [
        ("4", "24", "amostra oficial"),
        ("1", "1", "menor valor comum"),
        ("2", "2", "valor pequeno"),
        ("7", "5040", "valor intermediário"),
        ("10", "3628800", "valor alto"),
        ("12", "479001600", "valor máximo permitido"),
    ],
    "1828": [
        (
            "3\ntesoura papel\npedra Spock\nlagarto lagarto",
            "Caso #1: Bazinga!\nCaso #2: Raj trapaceou!\nCaso #3: De novo!",
            "amostra oficial",
        ),
        (
            "1\nSpock pedra",
            "Caso #1: Bazinga!",
            "caso único — a numeração começa em 1",
        ),
        (
            "5\nSpock Spock\nSpock tesoura\ntesoura Spock\nSpock papel\npapel Spock",
            "Caso #1: De novo!\nCaso #2: Bazinga!\nCaso #3: Raj trapaceou!\n"
            "Caso #4: Raj trapaceou!\nCaso #5: Bazinga!",
            "Spock escrito com S maiúsculo, como vem na entrada deste problema",
        ),
        (
            "6\npedra tesoura\npedra lagarto\nlagarto papel\nlagarto Spock\n"
            "papel pedra\ntesoura lagarto",
            "Caso #1: Bazinga!\nCaso #2: Bazinga!\nCaso #3: Bazinga!\n"
            "Caso #4: Bazinga!\nCaso #5: Bazinga!\nCaso #6: Bazinga!",
            "Sheldon vence todas — detecta matriz de regras invertida",
        ),
    ],
    "1873": [
        (
            "6\npedra tesoura\nspock spock\npapel lagarto\nlagarto spock\n"
            "tesoura pedra\npapel spock",
            "rajesh\nempate\nsheldon\nrajesh\nsheldon\nrajesh",
            "casos variados",
        ),
        (
            "5\npedra pedra\npapel papel\ntesoura tesoura\nlagarto lagarto\nspock spock",
            "empate\nempate\nempate\nempate\nempate",
            "todos empates",
        ),
        (
            "4\nspock tesoura\nspock pedra\ntesoura papel\ntesoura lagarto",
            "rajesh\nrajesh\nrajesh\nrajesh",
            "rajesh vence todas — detecta matriz de regras invertida",
        ),
        (
            "4\ntesoura spock\npedra spock\npapel tesoura\nlagarto tesoura",
            "sheldon\nsheldon\nsheldon\nsheldon",
            "sheldon vence todas — detecta ordem de leitura invertida",
        ),
    ],
}


def formatar_casos_teste(problema_id):
    """Monta o bloco de casos de teste que vai dentro do prompt.

    Devolve string vazia se o problema não tiver casos cadastrados, para que o
    prompt continue funcionando com problemas novos ainda sem bateria de testes.
    """
    casos = casos_teste.get(str(problema_id))
    if not casos:
        return ""

    linhas = []
    for i, (entrada, saida, descricao) in enumerate(casos, 1):
        linhas.append(
            f"Caso {i} ({descricao})\n"
            f"  Entrada:\n{_indentar(entrada)}\n"
            f"  Saída esperada:\n{_indentar(saida)}"
        )
    return "\n\n".join(linhas)


def _indentar(texto, espacos=4):
    prefixo = " " * espacos
    return "\n".join(prefixo + linha for linha in str(texto).split("\n"))


def montar_prompt(problem_description, student_code, problema_id=None):
    casos_teste_formatados = formatar_casos_teste(problema_id) or (
        "(nenhum caso de teste cadastrado para este problema — avalie apenas "
        "pela descrição acima)"
    )
    return f"""
Você é um professor de ciência da computação experiente, especialista em pedagogia para iniciantes em programação.
Sua tarefa é analisar o código de um aluno para um problema específico e fornecer um feedback formativo, utilizando a metodologia "Logic-of-Thought" (LoT).

Regras INEGOCIÁVEIS:
1. NUNCA forneça a solução correta. Isso significa: NÃO escreva o código corrigido e NÃO descreva o algoritmo em texto plano (Ex: É PROIBIDO dizer "adicione +1 no range" ou "inverta a ordem das variáveis" ou "coloque dois pontos no final do for").
2. Foco em UM ERRO POR VEZ (Dica Progressiva). Se houver erro de sintaxe, foque APENAS no erro de sintaxe. Não mencione erros de lógica até que o código do aluno consiga rodar.
3. Método Socrático Estrito: A sua dica de ação deve ser uma pergunta guiada ou um CASO DE TESTE (Ex: "Simule seu código no papel com X=20 e Y=10. Quantas vezes o laço executa?").
4. O código deve ser avaliado ESTRITAMENTE sob as regras de sintaxe da linguagem Python 3. Ignorar comentários iniciados com '#'.
5. Use os CASOS DE TESTE fornecidos como a fonte da verdade sobre o comportamento do código.
   Só marque is_correct como true se o código produzir a saída esperada em TODOS eles.
   Se algum caso falhar, o erro é real: não classifique o código como correto.
6. Ao citar um caso de teste na dica, você PODE mostrar a entrada e perguntar o que o
   código do aluno produz com ela. Você NÃO PODE revelar a saída esperada, nem dizer que
   a saída do aluno está diferente dela — isso entregaria a correção e viola a regra 1.

[Descrição do Problema]
{problem_description}

[Casos de Teste Oficiais]
Estes casos definem o comportamento correto. Simule o código do aluno mentalmente em cada
um deles antes de responder. Eles são para o SEU uso na verificação; as saídas esperadas
NUNCA devem aparecer no feedback_formativo nem na dica_acao.
{casos_teste_formatados}

[Código do Aluno]
{student_code}

Na fase_3_traducao, planeje a abordagem, mas não escreva a solução aqui também.
Retorne ESTRITAMENTE no formato JSON abaixo:
{{
  "logic_of_thought": {{
    "fase_1_extracao": "<Atue como um interpretador Python 3. Leia o código e identifique a árvore lógica que o aluno implementou.>",
    "fase_2_extensao": "<Percorra os casos de teste um a um dizendo, para cada um, qual saída o código do aluno produziria e se ela bate com a esperada. Liste então todos os erros encontrados e ELEJA APENAS O ERRO MAIS CRÍTICO para focar agora (Sintaxe > Erro de Execução > Erro de Lógica).>",
    "fase_3_traducao": "<Planeje a abordagem pedagógica para este ÚNICO erro. Prefira usar a entrada de um dos casos de teste que falharam: qual pergunta socrática fará o aluno perceber a falha sozinho?>"
  }},
  "is_correct": <true ou false>,
  "linha_do_erro": <numero_da_linha_do_erro_escolhido_na_fase_2>,
  "tipo_de_erro": "<Sintaxe, Execução ou Lógica>",
  "feedback_formativo": "<Um comentário curto e amigável dizendo vagamente onde a execução tropeça. Fale de APENAS UM erro.>",
  "dica_acao": "<Uma pergunta socrática ou um caso de teste para o aluno investigar. NUNCA diga o que está faltando.>",
  "casos_que_falham": [<lista com os números dos casos de teste que o código não passa; lista vazia se passar em todos>]
}}
"""


def chamar_gemini(problem_description, student_code, problema_id=None):
    prompt = montar_prompt(problem_description, student_code, problema_id)
    response = client.models.generate_content(
        model=MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.2,
        ),
    )
    return json.loads(response.text)


def avaliar(problema_id, student_code):
    """Avalia o código de um aluno para um problema do banco.

    Levanta ValueError se o problema não existir ou o código estiver vazio.
    """
    problem_description = banco_problemas.get(str(problema_id))
    if not problem_description:
        raise ValueError(f"Problema {problema_id} não encontrado no banco de problemas.")

    if not student_code or not student_code.strip():
        raise ValueError("O código enviado está vazio.")

    return chamar_gemini(problem_description, student_code, str(problema_id))