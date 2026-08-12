"""
Aplicação web de feedback formativo para alunos de Python.

O aluno escolhe um problema, escreva seu código e recebe uma mensagem de
feedback junto de uma dica socrática, geradas pelo mesmo prompt
Logic-of-Thought usado no script de avaliação em lote.

Rodar:
    python app.py
Depois abrir http://127.0.0.1:5000
"""
from flask import Flask, jsonify, render_template, request

from avaliador import avaliar, banco_problemas, titulos_problemas

app = Flask(__name__)


@app.route("/")
def index():
    problemas = [
        {
            "id": pid,
            "titulo": titulos_problemas.get(pid, pid),
            "descricao": descricao,
        }
        for pid, descricao in banco_problemas.items()
    ]
    return render_template("index.html", problemas=problemas)


@app.route("/api/avaliar", methods=["POST"])
def api_avaliar():
    dados = request.get_json(silent=True) or {}
    problema_id = dados.get("problema")
    codigo = dados.get("codigo", "")
    provedor = dados.get("provedor", "gemini")

    try:
        resultado = avaliar(problema_id, codigo, provedor)
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400
    except Exception as e:  # falha de rede, API ou JSON inválido
        return jsonify({"erro": f"Falha ao consultar o avaliador: {e}"}), 502

    return jsonify({
        "is_correct": resultado.get("is_correct"),
        "linha_do_erro": resultado.get("linha_do_erro"),
        "tipo_de_erro": resultado.get("tipo_de_erro"),
        "feedback_formativo": resultado.get("feedback_formativo"),
        "dica_acao": resultado.get("dica_acao"),
    })


if __name__ == "__main__":
    app.run(debug=True)
