"""
estruturar.py
Etapa 1 — estrutura o Excel bruto em conversas JSON (sem IA, sem custo)
Variáveis de ambiente: nenhuma
"""

import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


# ─────────────────────────────────────────────
# LEITURA DO EXCEL
# ─────────────────────────────────────────────

def ler_excel(caminho):
    xls = pd.read_excel(caminho, sheet_name=None)
    nome_aba = list(xls.keys())[0]
    df = xls[nome_aba]
    df = df[df.iloc[:, 1] != df.columns[1]].copy()
    df.columns = [
        'conta_nome', 'canal_chave', 'plataforma', 'contato_nome',
        'contato_tel', 'contato_instagram', 'contato_email',
        'msg_id', 'msg_data', 'msg_quem_enviou', 'msg_conteudo', 'conversa_url'
    ]
    df['data'] = pd.to_datetime(df['msg_data'], errors='coerce')
    df['session_id'] = df['conversa_url'].str.extract(r'id=([a-f0-9-]+)')

    def extrai_remetente(texto):
        m = re.match(r'De:\s*(.+?)\s*Para:', str(texto))
        return m.group(1).strip() if m else str(texto)

    df['remetente'] = df['msg_quem_enviou'].apply(extrai_remetente)
    return df.sort_values(['session_id', 'data'])


# ─────────────────────────────────────────────
# FUSÃO DE SESSÕES DO MESMO PACIENTE
# ─────────────────────────────────────────────

def fusao_sessoes(df, janela_dias=3):
    grupos = []
    for tel, grp_tel in df.groupby('contato_tel', dropna=False):
        if pd.isna(tel) or str(tel).strip() == '':
            for sid, grp_sid in grp_tel.groupby('session_id'):
                grupos.append(grp_sid)
            continue

        sessoes = []
        for sid, grp_sid in grp_tel.groupby('session_id'):
            sessoes.append((grp_sid['data'].min(), sid, grp_sid))
        sessoes.sort(key=lambda x: x[0])

        grupo_atual = sessoes[0][2].copy()
        data_fim_atual = sessoes[0][2]['data'].max()

        for i in range(1, len(sessoes)):
            data_inicio_prox, _, grp_prox = sessoes[i]
            delta = (data_inicio_prox - data_fim_atual).total_seconds() / 86400
            if delta <= janela_dias:
                grupo_atual = pd.concat([grupo_atual, grp_prox])
                data_fim_atual = max(data_fim_atual, grp_prox['data'].max())
            else:
                grupos.append(grupo_atual)
                grupo_atual = grp_prox.copy()
                data_fim_atual = grp_prox['data'].max()
        grupos.append(grupo_atual)
    return grupos


# ─────────────────────────────────────────────
# MONTAGEM DO TEXTO DA CONVERSA
# ─────────────────────────────────────────────

def montar_texto(grp):
    grp = grp.sort_values('data')
    linhas = []
    for _, row in grp.iterrows():
        dt = row['data'].strftime('%d/%m/%Y %H:%M') if pd.notna(row['data']) else ''
        remetente = row['remetente'] if pd.notna(row['remetente']) else ''
        conteudo = str(row['msg_conteudo']) if pd.notna(row['msg_conteudo']) else ''
        if conteudo.startswith('[Imagem]'): conteudo = '[imagem]'
        elif conteudo.startswith('[Documento]'): conteudo = '[documento]'
        elif conteudo.startswith('[Áudio]') or conteudo.startswith('[Audio]'): conteudo = '[áudio]'
        elif conteudo.startswith('[Vídeo]') or conteudo.startswith('[Video]'): conteudo = '[vídeo]'
        elif conteudo.startswith('*Atenção:'): conteudo = '[mensagem não suportada]'
        elif len(conteudo) > 800: conteudo = conteudo[:800] + '...'
        linhas.append(f'[{dt}] {remetente}: {conteudo}')
    return '\n'.join(linhas)


# ─────────────────────────────────────────────
# PRÉ-CLASSIFICAÇÃO (sem IA)
# ─────────────────────────────────────────────

INTERNOS = [
    'recepção bio campos', 'bio campos', 'cirurgias eletivas', 'emsella',
    'tallyta', 'emsellamogi', 'builderwise', 'fasdesigner', 'a2 systems',
    'ecoville', 'suporte -', 'climes clínica médica'
]


def pre_classificar(grp, texto):
    nome = str(grp['contato_nome'].iloc[0]).lower() if pd.notna(grp['contato_nome'].iloc[0]) else ''
    plataforma = str(grp['plataforma'].iloc[0]).lower() if pd.notna(grp['plataforma'].iloc[0]) else ''
    t = texto.lower()

    if any(k in nome for k in INTERNOS):
        return 'interno_fornecedor'

    if 'instagram' in plataforma:
        palavras = ['consulta', 'agend', 'procedimento', 'cirurgia', 'exame',
                    'vasectomia', 'implanon', 'diu', 'ginecolog', 'médico']
        if not any(p in t for p in palavras):
            return 'instagram_sem_interesse'

    padroes_lembrete = [
        'essa mensagem é para confirmar sua consulta',
        'sua atendimento está confirmada para o dia',
        'sua consulta está confirmada',
    ]
    if any(p in t[:300] for p in padroes_lembrete):
        return 'lembrete_consulta'

    if len(grp) <= 2 and len(t) < 100:
        return 'conversa_fragmentada'

    return None


# ─────────────────────────────────────────────
# FUNÇÃO PRINCIPAL
# ─────────────────────────────────────────────

def estruturar(caminho_arquivo, log_func=print, janela_dias=3):
    log_func('📂 Lendo arquivo...')
    df = ler_excel(caminho_arquivo)
    log_func(f'📊 {len(df)} mensagens | {df["session_id"].nunique()} sessões originais')
    log_func(f'🔗 Fundindo sessões do mesmo paciente...')
    grupos = fusao_sessoes(df, janela_dias)
    log_func(f'✅ {len(grupos)} conversas após fusão')

    stats = {
        'total': len(grupos), 'requer_ia': 0,
        'lembrete_consulta': 0, 'interno': 0,
        'instagram': 0, 'fragmentada': 0
    }

    conversas = []
    for grp in grupos:
        grp = grp.sort_values('data')
        texto = montar_texto(grp)
        pre_class = pre_classificar(grp, texto)

        conv = {
            'session_ids': list(grp['session_id'].dropna().unique()),
            'paciente_nome': str(grp['contato_nome'].iloc[0]) if pd.notna(grp['contato_nome'].iloc[0]) else '',
            'paciente_tel': str(grp['contato_tel'].iloc[0]) if pd.notna(grp['contato_tel'].iloc[0]) else '',
            'plataforma': str(grp['plataforma'].iloc[0]) if pd.notna(grp['plataforma'].iloc[0]) else '',
            'data_inicio': grp['data'].min().strftime('%d/%m/%Y %H:%M') if pd.notna(grp['data'].min()) else '',
            'data_fim': grp['data'].max().strftime('%d/%m/%Y %H:%M') if pd.notna(grp['data'].max()) else '',
            'total_mensagens': len(grp),
            'pre_classificacao': pre_class,
            'requer_ia': pre_class is None,
            'conversa_texto': texto,
            'tipo_conversa': pre_class if pre_class else None,
            'medico': None, 'procedimento': None,
            'gerou_agendamento': None, 'motivo_nao_agendamento': None,
            'responsavel': None, 'resumo': None, 'evidencia': None,
            'analisado': pre_class is not None
        }

        if pre_class:
            key = pre_class.replace('interno_fornecedor', 'interno') \
                           .replace('instagram_sem_interesse', 'instagram') \
                           .replace('conversa_fragmentada', 'fragmentada')
            if key in stats:
                stats[key] += 1
        else:
            stats['requer_ia'] += 1

        conversas.append(conv)

    return conversas, stats
