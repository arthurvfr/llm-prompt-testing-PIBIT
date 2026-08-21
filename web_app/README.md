# Aplicação Web — Feedback Formativo de Python

Interface em que o aluno escolhe um problema, escreve seu código em Python e
recebe **uma mensagem de feedback** junto de **uma dica socrática**, geradas
pelo mesmo prompt Logic-of-Thought (LoT) usado em `../scripts/avaliar_lote_gemini.py`.

## Estrutura

```
web_app/
├── app.py              # servidor Flask (rotas / e /api/avaliar)
├── avaliador.py        # banco de problemas + prompt LoT + chamada ao Gemini/LLM local
├── templates/
│   └── index.html
└── static/
    ├── style.css
    └── script.js
```

As dependências ficam no `requirements.txt` único na raiz do repositório (veja
o README principal) — não há um `requirements.txt` próprio desta pasta.

## Fonte do feedback: Gemini ou LLM local

A interface tem um seletor "Fonte do feedback" com duas opções, enviadas em
`provedor` no POST de `/api/avaliar`:

- **Gemini (nuvem)** — padrão, usa a API do Gemini (precisa de `GEMINI_API_KEY`).
- **LLM local (llama-server)** — usa um servidor llama.cpp local, com a mesma
  lógica de `../scripts/avaliar_lote_local.py` (endpoint `http://localhost:1234/v1/chat/completions`).
  Suba o servidor local (ex.: `../iniciar_server.bat`) antes de escolher essa opção.

O prompt Logic-of-Thought e o banco de problemas/casos de teste são os mesmos
para as duas fontes — só muda quem responde.

## Chave da API

Usa `GEMINI_API_KEY` da variável de ambiente; se não existir, importa o
`config.py` da pasta raiz do projeto (mesmo comportamento do script de lote).
Só é necessária se você for usar a fonte "Gemini (nuvem)" — a `chave` é
carregada de forma preguiçosa, então rodar só com o LLM local funciona sem ela.

## Como rodar

```powershell
..\venv\Scripts\pip.exe install -r ..\requirements.txt
..\venv\Scripts\python.exe app.py
```

Abra <http://127.0.0.1:5000>.

## Observação pedagógica

O bloco `logic_of_thought` da resposta do modelo **não** é enviado ao navegador —
apenas `feedback_formativo`, `dica_acao`, `tipo_de_erro` e `linha_do_erro`.
Assim o raciocínio interno (que pode conter a solução) nunca chega ao aluno.
