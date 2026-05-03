"""
notify.py — Envia uma mensagem WhatsApp via CallMeBot para todos os destinos
configurados nas variáveis de ambiente.

Configuração (env vars):
    WHATSAPP_PHONE_1, WHATSAPP_APIKEY_1   # destino 1 (obrigatório)
    WHATSAPP_PHONE_2, WHATSAPP_APIKEY_2   # destino 2 (opcional)

Se um par estiver vazio (ex: ainda não autorizamos o telefone 2), é
silenciosamente ignorado. Um par incompleto (só phone, ou só apikey) é
considerado erro de configuração e gera aviso explícito mas não aborta.

Uso:
    python notify.py "sua mensagem aqui"
"""
from __future__ import annotations

import os
import sys
import urllib.parse

import requests

CALLMEBOT_URL = 'https://api.callmebot.com/whatsapp.php'
TIMEOUT_S = 30


def enviar(phone: str, apikey: str, msg: str) -> tuple[bool, str]:
    """Tenta enviar uma mensagem. Retorna (ok, descricao)."""
    url = (
        f'{CALLMEBOT_URL}'
        f'?phone={urllib.parse.quote(phone)}'
        f'&text={urllib.parse.quote(msg)}'
        f'&apikey={urllib.parse.quote(apikey)}'
    )
    try:
        r = requests.get(url, timeout=TIMEOUT_S)
    except requests.RequestException as e:
        return False, f'erro de rede: {e}'

    # CallMeBot retorna 200 com HTML em sucesso, e 203 em alguns casos de
    # rate limit / erro lógico. Tratamos 2xx como sucesso aparente.
    if 200 <= r.status_code < 300:
        return True, f'HTTP {r.status_code}'
    return False, f'HTTP {r.status_code} — {r.text[:200]}'


def main(msg: str) -> int:
    if not msg.strip():
        print('❌ Mensagem vazia', file=sys.stderr)
        return 1

    enviados = 0
    falhas = 0

    for i in (1, 2):
        phone = os.environ.get(f'WHATSAPP_PHONE_{i}', '').strip()
        apikey = os.environ.get(f'WHATSAPP_APIKEY_{i}', '').strip()

        # Par totalmente vazio → ignora (destino opcional)
        if not phone and not apikey:
            continue

        # Par parcial → erro de config, mas não aborta
        if not phone or not apikey:
            print(f'⚠️  Destino {i}: configuração incompleta '
                  f'(phone={"set" if phone else "vazio"}, '
                  f'apikey={"set" if apikey else "vazio"}) — pulando')
            falhas += 1
            continue

        ok, info = enviar(phone, apikey, msg)
        # Mascara o telefone no log: +5511******0050
        masked = phone[:5] + '*' * max(0, len(phone) - 9) + phone[-4:]
        if ok:
            print(f'✅ Destino {i} ({masked}): {info}')
            enviados += 1
        else:
            print(f'❌ Destino {i} ({masked}): {info}')
            falhas += 1

    print(f'\nResumo: {enviados} enviado(s), {falhas} falha(s)')

    # Falha apenas se NINGUÉM recebeu — destino 2 ausente é cenário esperado
    return 0 if enviados > 0 else 1


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit('uso: python notify.py "mensagem"')
    sys.exit(main(sys.argv[1]))
