# Aplicação Web — Feedback Formativo de Python

Interface em que o aluno escolhe um problema, escreve seu código em Python e
recebe **uma mensagem de feedback** junto de **uma dica socrática**, geradas
pelo mesmo prompt Logic-of-Thought (LoT) usado em `../avaliar_lote.py`.

## Estrutura

```
web_app/
├── app.py              # servidor Flask (rotas / e /api/avaliar)
├── avaliador.py        # banco de problemas + prompt LoT + chamada ao Gemini
├── requirements.txt
├── templates/
│   └── index.html
└── static/
    ├── style.css
    └── script.js
```

## Chave da API

Usa `GEMINI_API_KEY` da variável de ambiente; se não existir, importa o
`config.py` da pasta raiz do projeto (mesmo comportamento do script de lote).

## Como rodar

```powershell
..\venv\Scripts\pip.exe install -r requirements.txt
..\venv\Scripts\python.exe app.py
```

Abra <http://127.0.0.1:5000>.

## Observação pedagógica

O bloco `logic_of_thought` da resposta do modelo **não** é enviado ao navegador —
apenas `feedback_formativo`, `dica_acao`, `tipo_de_erro` e `linha_do_erro`.
Assim o raciocínio interno (que pode conter a solução) nunca chega ao aluno.
