# 🏥 ClinicDesk — Análise de Conversas

Sistema web para análise automática de conversas de atendimento da **Clínica Climes**, desenvolvido em Python + Flask com integração ao Claude AI (Anthropic).

---

## 📋 O que o sistema faz

1. **Etapa 1 — Estruturação (sem custo):** lê o arquivo Excel exportado do sistema da clínica, funde sessões fragmentadas do mesmo paciente, pré-classifica lembretes, internos e Instagram automaticamente.

2. **Etapa 2 — Análise com IA:** envia as conversas ao Claude API, que classifica cada uma como agendamento confirmado ou não, identifica o médico, procedimento e motivo de não agendamento.

3. **Output:** gera um Excel com duas abas — conversas completas e análise executiva com KPIs, motivos de não agendamento, agendamentos por médico e falhas da clínica.

---

## 🚀 Como usar

1. Acesse o sistema pelo navegador
2. Clique em **Novo Arquivo**
3. Suba o `.xlsx` exportado da plataforma da clínica
4. Aguarde a estruturação (Etapa 1 — instantânea)
5. Clique em **Analisar com IA** (Etapa 2)
6. Baixe o Excel gerado

---

## 🏗️ Estrutura do projeto

```
clinica_app/
├── app.py              # Servidor web Flask
├── estruturar.py       # Etapa 1 — estruturação sem IA
├── analisar.py         # Etapa 2 — análise com Claude API
├── aprendizado.json    # Memória acumulativa de erros e acertos
├── requirements.txt    # Dependências Python
├── templates/
│   ├── index.html      # Dashboard principal
│   └── processar.html  # Upload e progresso
└── static/
    └── style.css       # Visual da interface
```

---

## ⚙️ Variáveis de ambiente

| Variável | Descrição |
|---|---|
| `ANTHROPIC_API_KEY` | Chave de API do Claude (Anthropic) |
| `SECRET_KEY` | Chave secreta do Flask (opcional) |

---

## 📦 Instalação local

```bash
pip install -r requirements.txt
python app.py
```

---

## 🧠 Sistema de aprendizado

O arquivo `aprendizado.json` acumula exemplos de erros corrigidos em auditorias anteriores. A cada revisão, rode:

```bash
python analisar.py --aprender arquivo_revisao.xlsx
```

O Claude incorpora esses exemplos automaticamente nas próximas análises, melhorando a precisão a cada rodada.

---

## 💰 Custo estimado

- ~US$ 0,02 por conversa analisada
- Pré-classificação elimina lembretes, internos e Instagram antes de chamar a IA
- Custo típico por rodada semanal: US$ 8–12

---

## 🏥 Clínica Climes — Mogi das Cruzes/SP
