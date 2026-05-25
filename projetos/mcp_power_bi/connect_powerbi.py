#!/usr/bin/env python3
"""
Detecta a porta do Analysis Services do Power BI Desktop em execução
e conecta ao modelo via pyadomd para executar uma consulta DAX simples.

Uso:
  1. Instale dependências: pip install -r requirements.txt
  2. Abra o Power BI Desktop com o relatório desejado
  3. Execute: python connect_powerbi.py
"""
import sys
import time
try:
    import psutil
except Exception:
    print("Erro: instale 'psutil' (pip install psutil)")
    raise

try:
    from pyadomd import Pyadomd
except Exception:
    print("Erro: instale 'pyadomd' (pip install pyadomd). Requer ADOMD.NET no Windows.")
    raise


def find_msmdsrv_port():
    """Procura processos 'msmdsrv.exe' e retorna a primeira porta TCP em LISTEN."""
    for proc in psutil.process_iter(['pid', 'name']):
        name = (proc.info.get('name') or '').lower()
        if 'msmdsrv' in name:
            pid = proc.info['pid']
            try:
                conns = psutil.net_connections(kind='inet')
            except Exception:
                conns = []
            for c in conns:
                if c.pid == pid and c.status.upper().startswith('LIST'):
                    laddr = c.laddr
                    # laddr pode ser tuple (ip, port) ou psutil._common.addr
                    try:
                        if hasattr(laddr, 'port'):
                            return laddr.port
                        elif isinstance(laddr, tuple) and len(laddr) >= 2:
                            return laddr[1]
                    except Exception:
                        pass
    return None


def connect_and_query(port):
    conn_str = f"Data Source=localhost:{port};Initial Catalog=Model;Integrated Security=SSPI"
    print(f"Tentando conectar: {conn_str}")
    try:
        with Pyadomd(conn_str) as conn:
            conn.open()
            cur = conn.cursor()
            query = 'EVALUATE ROW("Ping", 1)'
            cur.execute(query)
            rows = cur.fetchall()
            print('Resultado da consulta:')
            for r in rows:
                print(r)
    except Exception as e:
        print('Falha ao conectar/executar consulta:')
        print(e)


def main():
    print('Procurando processo do Analysis Services (msmdsrv)...')
    port = find_msmdsrv_port()
    if not port:
        print('Não encontrei msmdsrv em execução. Verifique se o Power BI Desktop está aberto com um relatório.')
        sys.exit(1)
    print(f'Encontrado msmdsrv ouvindo na porta: {port}')
    connect_and_query(port)


if __name__ == '__main__':
    main()
