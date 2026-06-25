"""Envia os 3 DOCX da SONP 80/2026 (Pleno, 1ª Câmara, 2ª Câmara) em UM e-mail
via Outlook (pywin32), para raphael.goncalves e glaucia.calvet.
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import win32com.client as win32

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"

ANEXOS = [
    OUTPUT / "PAUTA_SONP_80_PLENO_2026.docx",
    OUTPUT / "PAUTA_SONP_80_1CAMARA_2026.docx",
    OUTPUT / "PAUTA_SONP_80_2CAMARA_2026.docx",
]

DESTINATARIOS = [
    "raphael.goncalves@tcmsp.tc.br",
    "glaucia.calvet@tcmsp.tc.br",
]

ACCOUNT_HINT = os.getenv("TCM_EMAIL_ACCOUNT", "pautaeletronica@tcmsp.tc.br")

ASSUNTO = "Pauta SONP 80/2026 - Pleno, 1ª e 2ª Câmara (revisão Glaucia, 3ª rodada)"

CORPO_HTML = """
<p>Prezados Rapha e Glaucia,</p>

<p>Seguem em anexo as 3 pautas da <b>SONP 80/2026</b> (Pleno, 1ª Câmara e 2ª Câmara),
regeradas com a 2ª rodada de correções da Glaucia.</p>

<p><b>Correções aplicadas nesta rodada:</b></p>
<ul>
  <li>Título do Pleno: "SESSÃO ORDINÁRIA <b>NÃO PRESENCIAL EM AMBIENTE
      VIRTUAL DO TRIBUNAL PLENO</b> DO TRIBUNAL DE CONTAS";</li>
  <li>Após "- II - JULGAMENTOS", cabeçalho da competência
      (<b>"PROCESSOS DO PLENO"</b> / <b>"PROCESSOS DA 1ª CÂMARA"</b> /
      <b>"PROCESSOS DA 2ª CÂMARA"</b>), <u>sublinhado</u> e em <span style="color:#800080">roxo</span>;</li>
  <li>Câmaras: linha do <b>Presidente da Câmara</b> em negrito (1ª: Conselheiro
      Presidente Domingos Dissei; 2ª: <b>Conselheiro Ricardo Torres</b> &mdash; sem
      "Vice-Presidente", conforme padrão da chefe);</li>
  <li>DD como Relator aparece sempre como
      "<b>I - CONSELHEIRO PRESIDENTE DOMINGOS DISSEI, na qualidade de Relator</b>"
      (tanto no Pleno quanto na 1ª Câmara, com o sufixo em caixa mista);</li>
  <li>Câmaras: Relatores sem processos saem com
      "<b>(Sem processos para relatar)</b>" (inclui a Conselheira Substituta no
      lugar do RB na 1ª Câmara);</li>
  <li>Conselheira Substituta com concordância feminina:
      "<b>RELATORA CONSELHEIRA SUBSTITUTA DANIELA FARIAS</b>";</li>
  <li>Recurso / Embargos / Representação / Acompanhamento: apenas o tipo
      principal sai em negrito - o subtipo do processo julgado fica sem negrito
      (corrige TC/003490/2022 e os "Acompanhamento - Execução Contratual");</li>
  <li>Bloco final na ordem correta: <i>"São Paulo, &lt;data&gt;."</i> &rarr; linha em
      branco &rarr; <b>ROSELI DE MORAIS CHAVES</b> &rarr; "Subsecretária-Geral".
      A data usada é a do disparo do e-mail.</li>
</ul>

<p><b>Mantido das rodadas anteriores:</b> Resolução nº 24/2025 + Instrução nº 01/2025;
numeração contínua por Relator; ordenação por espécie/ano/número; siglas de
Conselheiros removidas do objeto; advogados em linha própria (fonte 10pt); "Verificado
até a peça..." e "Retorno à pauta..." removidos; "Tramitam em conjunto os TCs..."
no plural; aspas tipográficas normalizadas; substitutos no final no Pleno.</p>

<p><b>Pontos que dependem do cadastro no e-TCM</b> (a automação reflete o que estiver
no objeto - precisam ser corrigidos pela Soninha antes da captura):</p>

<ul>
  <li><b>TC/000099/2017</b>: objeto com erros de digitação;</li>
  <li><b>TC/013380/2017</b>: falta a sigla do Procurador (FCCF);</li>
  <li><b>TC/003490/2022</b>: aparece "OBJETO" sobrando após o nome da Concessionária;</li>
  <li><b>TC/006201/2023</b>: texto extra entre a sigla do Procurador e os advogados;</li>
  <li><b>TC/002101/2025</b>: falta "edital de" antes de "Pregão Eletrônico";</li>
  <li><b>TC/006615/2019 / TC/006616/2019</b>: itens englobados (7 e 8) e tramitação
      conjunta não vieram do cadastro; advogado e sigla do Procurador faltando;
      objeto com "?" residual de en-dash colado de fora;</li>
  <li><b>TC/004042/2022</b>: falta crase em "ligadas à Rede" e advogada não cadastrada;</li>
  <li><b>TC/004918/2018</b>: falta acento em "CONSÓRCIO" e advogados não cadastrados;</li>
  <li><b>TC/011038/2018</b>: texto do objeto desatualizado em relação ao que a Soninha
      definiu (correção pendente no e-TCM);</li>
  <li><b>TC/003265/2018 e Representações em tramitação conjunta</b>: faltam
      Representante / Representado / advogados no cadastro;</li>
  <li><b>TC/017955/2024</b>: texto do objeto incompleto/sintaticamente fragmentado
      (escolha da chefe).</li>
</ul>

<p><b>Limitação conhecida da automação (próxima rodada):</b> a regra para evitar a
dupla RT/DD (oito processos da revisaria do RT que foram substituídos pela Daniela)
ainda não está modelada - hoje o robô reflete o revisor que vem do e-TCM. Vou subir
essa regra na próxima rodada.</p>

<p>Recomendação: para reduzir o vai-e-vem, dispararmos a captura no e-TCM
<b>após às 23h59 da sexta-feira</b>, quando a Soninha já finalizou as correções.</p>

<p>Qualquer ajuste adicional, é só avisar.</p>

<p>Abraços,<br/>
Rapha<br/>
<small>Gerado automaticamente em {ts} pela automação Pauta TCM.</small></p>
"""


def _find_account(ns, hint: str):
    if not hint:
        return None
    h = hint.lower()
    try:
        for acc in ns.Session.Accounts:
            name = str(getattr(acc, "DisplayName", "") or "").lower()
            smtp = str(getattr(acc, "SmtpAddress", "") or getattr(acc, "Address", "") or "").lower()
            if h in name or h in smtp:
                return acc
    except Exception:
        pass
    return None


def main() -> None:
    # Valida anexos
    for p in ANEXOS:
        if not p.exists():
            raise SystemExit(f"Anexo nao encontrado: {p}")

    app = win32.gencache.EnsureDispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")

    mail = app.CreateItem(0)  # olMailItem
    mail.Subject = ASSUNTO
    mail.HTMLBody = CORPO_HTML.format(
        ts=datetime.now().strftime("%d/%m/%Y %H:%M")
    )

    # Destinatarios (TO)
    for addr in DESTINATARIOS:
        r = mail.Recipients.Add(addr)
        try:
            r.Type = 1  # olTo
        except Exception:
            pass

    # Conta de envio
    acc = _find_account(ns, ACCOUNT_HINT)
    if acc is not None:
        try:
            mail.SendUsingAccount = acc
        except Exception:
            try:
                mail._oleobj_.Invoke(64209, 0, 8, 0, acc)
            except Exception as e:
                print(f"[email] Aviso: nao foi possivel fixar SendUsingAccount: {e}")
        try:
            mail.SentOnBehalfOfName = ACCOUNT_HINT
        except Exception:
            pass

    # Anexos (3)
    for p in ANEXOS:
        mail.Attachments.Add(str(p.resolve()))
        print(f"[email] Anexado: {p.name}")

    # Resolve destinatarios
    try:
        ok = bool(mail.Recipients.ResolveAll())
        print(f"[email] ResolveAll: {ok}")
    except Exception as e:
        print(f"[email] Aviso ResolveAll: {e}")

    # Envia
    mail.Send()
    print("[email] Enviado com sucesso.")


if __name__ == "__main__":
    main()
