"""
drive_sync.py — Helper para sincronizar arquivos entre o Google Drive e o
runner do GitHub Actions, usado pelo dashboard Bigtex.

Autenticação via Service Account. Credencial e IDs das pastas vêm de variáveis
de ambiente (em produção, são Secrets do GitHub Actions):

    GOOGLE_SERVICE_ACCOUNT_JSON  Conteúdo (string) do JSON da Service Account
    DRIVE_INBOX_ID               ID da pasta INBOX no Drive
    DRIVE_BASES_ID               ID da pasta BASES no Drive

Subcomandos:

    drive_sync.py download-inbox --dest ./inbox [--names-out ./inbox_files.txt]
    drive_sync.py download-bases --dest ./bases
    drive_sync.py upload-bases   --src  ./bases
    drive_sync.py clear-inbox    --names a.xlsx b.xlsx
    drive_sync.py clear-inbox    --names-file ./inbox_files.txt

Notas sobre a arquitetura (Solução A — Drive pessoal, sem Workspace):

    Service Accounts em Drive pessoal NÃO têm cota de armazenamento, então
    NÃO conseguem CRIAR arquivos. Por isso:

      • upload-bases SÓ FAZ UPDATE (sobrescreve arquivo já existente).
        Se um arquivo local não tem correspondente no Drive, é REGISTRADO
        como "faltante" e o comando termina com exit_code=1, mas os demais
        ainda são processados.

      • Pastas PROCESSADOS e LOGS NÃO são usadas — esses artefatos vão pro
        Git (versionados via commits do workflow).

    No setup inicial, o usuário (proprietário do Drive) sobe manualmente os
    arquivos canônicos pra BASES. Depois, o workflow só faz UPDATE neles.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive']

# Sanity check no tamanho dos arquivos (planilhas/PDFs do projeto não chegam perto).
MAX_FILE_BYTES = 100 * 1024 * 1024


# ----------------------------------------------------------------------------
# Setup
# ----------------------------------------------------------------------------

def _get_env(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        sys.exit(f'❌ Variável de ambiente obrigatória não definida: {name}')
    return val


def get_service():
    """Constrói o cliente da Drive API com credenciais de Service Account.

    Aceita o JSON via duas formas:
      1) GOOGLE_SERVICE_ACCOUNT_JSON  conteúdo cru do JSON (uso em CI)
      2) GOOGLE_APPLICATION_CREDENTIALS  caminho pra arquivo (uso local)
    """
    raw = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if raw:
        info = json.loads(raw)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES
        )
    else:
        path = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        if not path:
            sys.exit(
                '❌ Defina GOOGLE_SERVICE_ACCOUNT_JSON (conteúdo) ou '
                'GOOGLE_APPLICATION_CREDENTIALS (caminho do arquivo).'
            )
        creds = service_account.Credentials.from_service_account_file(
            path, scopes=SCOPES
        )

    return build('drive', 'v3', credentials=creds, cache_discovery=False)


# ----------------------------------------------------------------------------
# Operações primitivas no Drive
# ----------------------------------------------------------------------------

def listar_arquivos(service, folder_id: str) -> list[dict]:
    """Lista todos os arquivos (não-pastas, não-trashed) de uma pasta."""
    arquivos = []
    page_token = None
    query = (
        f"'{folder_id}' in parents and "
        "mimeType != 'application/vnd.google-apps.folder' and "
        "trashed = false"
    )
    while True:
        resp = service.files().list(
            q=query,
            fields='nextPageToken, files(id, name, mimeType, size, md5Checksum)',
            pageSize=1000,
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        arquivos.extend(resp.get('files', []))
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    return arquivos


def baixar_arquivo(service, file_id: str, destino: Path) -> None:
    """Baixa um arquivo do Drive pra um caminho local."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    request = service.files().get_media(
        fileId=file_id, supportsAllDrives=True
    )
    buf = io.FileIO(str(destino), 'wb')
    try:
        downloader = MediaIoBaseDownload(buf, request, chunksize=1024 * 1024)
        done = False
        while not done:
            _status, done = downloader.next_chunk()
    finally:
        buf.close()


def atualizar_arquivo(service, file_id: str, src: Path) -> dict:
    """Atualiza o conteúdo de um arquivo existente no Drive (UPDATE)."""
    media = MediaFileUpload(str(src), resumable=False)
    return service.files().update(
        fileId=file_id,
        media_body=media,
        supportsAllDrives=True,
        fields='id, name',
    ).execute()


def remover_de_pasta(service, file_id: str, folder_id: str) -> None:
    """Desvincula um arquivo de uma pasta (NÃO deleta o arquivo).

    Usa `removeParents` em vez de mover pra lixeira porque a Service Account,
    sendo apenas Editora da pasta (não proprietária dos arquivos dentro),
    não tem permissão para deletar/trashear arquivos do usuário. Mas pode
    desvinculá-los da pasta — isso é uma operação na pasta, que ela tem
    permissão de editar.

    O arquivo desvinculado continua existindo no "Meu Drive" do dono
    (propriedade preservada). Cópia auditável fica em processados/<data>/
    no Git, então órfãos no Drive podem ser apagados pelo usuário sem perda.
    """
    service.files().update(
        fileId=file_id,
        removeParents=folder_id,
        supportsAllDrives=True,
        fields='id, parents',
    ).execute()


# ----------------------------------------------------------------------------
# Subcomandos
# ----------------------------------------------------------------------------

def cmd_download_inbox(args):
    """Baixa todos os arquivos do INBOX pra ./inbox/.

    Lista nomes em --names-out (default: ./inbox_files.txt no diretório-pai
    do --dest), pra ser consumida depois por `clear-inbox --names-file`.
    """
    service = get_service()
    folder_id = _get_env('DRIVE_INBOX_ID')
    dest = Path(args.dest)
    dest.mkdir(parents=True, exist_ok=True)

    arquivos = listar_arquivos(service, folder_id)
    print(f'📥 INBOX: {len(arquivos)} arquivo(s) encontrado(s)')
    for arq in arquivos:
        nome = arq['name']
        destino = dest / nome
        print(f'   ↓ {nome} ({arq.get("size", "?")} bytes)')
        baixar_arquivo(service, arq['id'], destino)

    nomes_path = (
        Path(args.names_out) if args.names_out
        else (dest.parent / 'inbox_files.txt')
    )
    nomes_path.parent.mkdir(parents=True, exist_ok=True)
    nomes_path.write_text(
        '\n'.join(a['name'] for a in arquivos), encoding='utf-8'
    )
    print(f'   📄 nomes gravados em {nomes_path}')


def cmd_download_bases(args):
    """Baixa todos os arquivos do BASES pra ./bases/."""
    service = get_service()
    folder_id = _get_env('DRIVE_BASES_ID')
    dest = Path(args.dest)
    dest.mkdir(parents=True, exist_ok=True)

    arquivos = listar_arquivos(service, folder_id)
    print(f'📥 BASES: {len(arquivos)} arquivo(s) encontrado(s)')
    for arq in arquivos:
        nome = arq['name']
        destino = dest / nome
        print(f'   ↓ {nome}')
        baixar_arquivo(service, arq['id'], destino)


def cmd_upload_bases(args):
    """Sobe arquivos de ./bases/ pra Drive(BASES) — APENAS UPDATE.

    Se um arquivo local NÃO tem correspondente no Drive, é registrado como
    faltante (SA não pode criar). Demais são atualizados normalmente.
    Termina com exit 1 se houver algum faltante.
    """
    service = get_service()
    folder_id = _get_env('DRIVE_BASES_ID')
    src = Path(args.src)
    if not src.is_dir():
        sys.exit(f'❌ Pasta não existe: {src}')

    arquivos_locais = sorted(p for p in src.iterdir() if p.is_file())
    if not arquivos_locais:
        print(f'⚠️  ./bases/ vazio: {src}')
        return

    # Indexa o que já existe no Drive (por nome)
    arquivos_drive = listar_arquivos(service, folder_id)
    por_nome = {a['name']: a for a in arquivos_drive}

    print(f'📤 BASES: {len(arquivos_locais)} arquivo(s) local(is)')
    atualizados, faltantes, erros = [], [], []

    for arq in arquivos_locais:
        if arq.stat().st_size > MAX_FILE_BYTES:
            erros.append(f'{arq.name}: tamanho > {MAX_FILE_BYTES} bytes')
            continue

        existente = por_nome.get(arq.name)
        if not existente:
            faltantes.append(arq.name)
            print(f'   ⚠️  FALTA NO DRIVE: {arq.name}')
            continue

        try:
            atualizar_arquivo(service, existente['id'], arq)
            atualizados.append(arq.name)
            print(f'   🔁 {arq.name}')
        except HttpError as e:
            erros.append(f'{arq.name}: {e}')
            print(f'   ❌ {arq.name}: {e}')

    print()
    print(f'   ✅ atualizados:  {len(atualizados)}')
    if faltantes:
        print(f'   ⚠️  faltantes:    {len(faltantes)}')
        print('        (precisam ser criados manualmente em BASES no Drive '
              'pelo proprietário, pois a Service Account não pode criar '
              'arquivos em Drive pessoal)')
        for nome in faltantes:
            print(f'          - {nome}')
    if erros:
        print(f'   ❌ erros:        {len(erros)}')
        for msg in erros:
            print(f'          - {msg}')

    if faltantes or erros:
        sys.exit(1)


def cmd_clear_inbox(args):
    """Remove (lixeira) arquivos do Drive(INBOX) pelos nomes dados.

    Recebe --names ou --names-file (um nome por linha).
    """
    service = get_service()
    folder_id = _get_env('DRIVE_INBOX_ID')

    nomes: list[str] = []
    if args.names:
        nomes.extend(args.names)
    if args.names_file:
        p = Path(args.names_file)
        if p.exists():
            nomes.extend(
                ln.strip() for ln in p.read_text(encoding='utf-8').splitlines()
                if ln.strip()
            )
    nomes = list(dict.fromkeys(nomes))  # dedup preservando ordem

    if not nomes:
        print('⚠️  Nenhum nome fornecido — nada a fazer')
        return

    print(f'🗑️  INBOX: removendo {len(nomes)} arquivo(s)')
    arquivos_inbox = listar_arquivos(service, folder_id)
    por_nome = {a['name']: a for a in arquivos_inbox}
    for nome in nomes:
        arq = por_nome.get(nome)
        if not arq:
            print(f'   ⚠️  não encontrado no INBOX: {nome}')
            continue
        try:
            remover_de_pasta(service, arq['id'], folder_id)
            print(f'   ✓ removido: {nome}')
        except HttpError as e:
            # Não interrompe o lote — loga e segue
            print(f'   ❌ falha ao remover {nome}: {e}')


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='Sincronização Drive ↔ runner para o dashboard Bigtex',
    )
    sub = parser.add_subparsers(dest='cmd', required=True)

    p = sub.add_parser('download-inbox', help='Baixa Drive(INBOX) → ./inbox/')
    p.add_argument('--dest', default='./inbox')
    p.add_argument('--names-out', default=None,
                   help='Caminho para gravar lista de nomes baixados '
                        '(default: ./inbox_files.txt no parent de --dest)')
    p.set_defaults(func=cmd_download_inbox)

    p = sub.add_parser('download-bases', help='Baixa Drive(BASES) → ./bases/')
    p.add_argument('--dest', default='./bases')
    p.set_defaults(func=cmd_download_bases)

    p = sub.add_parser('upload-bases', help='./bases/ → Drive(BASES) (apenas UPDATE)')
    p.add_argument('--src', default='./bases')
    p.set_defaults(func=cmd_upload_bases)

    p = sub.add_parser('clear-inbox',
                       help='Remove arquivos do Drive(INBOX) por nome')
    p.add_argument('--names', nargs='*', default=[],
                   help='Lista de nomes a remover')
    p.add_argument('--names-file', default=None,
                   help='Arquivo com um nome por linha (ex: ./inbox_files.txt)')
    p.set_defaults(func=cmd_clear_inbox)

    return parser


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except HttpError as e:
        sys.exit(f'❌ Erro Drive API: {e}')


if __name__ == '__main__':
    main()
