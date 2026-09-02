"""Consulta Gateway Status com autenticação delegada via Azure CLI.

Executação:
    Basta executar este arquivo diretamente pelo VS Code.

Requisitos:
    - Python 3.10+
    - Azure CLI instalado
    - Usuário com acesso aos Gateway Clusters no Power BI
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from pathlib import Path
from typing import Any


# ============================================================================
# CONFIGURAÇÕES
# ============================================================================

POWER_BI_RESOURCE = "https://analysis.windows.net/powerbi/api"

GATEWAY_STATUS_URL = (
    "https://api.powerbi.com/v2.0/myorg/gatewayclusters"
    "?$expand=permissions,memberGateways&$skip=0"
)

SCRIPT_DIR = Path(__file__).resolve().parent

AZURE_PROFILE_DIR = SCRIPT_DIR / ".azure_gateway_status"


# ============================================================================
# AZURE CLI
# ============================================================================

def azure_cli() -> str:
    """
    Localiza o executável do Azure CLI.

    No Windows normalmente será encontrado como az.cmd.
    """

    executable = (
        shutil.which("az.cmd")
        or shutil.which("az.exe")
        or shutil.which("az")
    )

    if not executable:
        raise RuntimeError(
            "Azure CLI não encontrado.\n\n"
            "Instale o Azure CLI antes de executar esta análise.\n"
            "Após a instalação, reinicie o VS Code."
        )

    return executable


def azure_environment() -> dict[str, str]:
    """
    Cria ambiente isolado para autenticação Azure.

    O Assessment Analyzer utilizará seu próprio diretório de sessão,
    sem interferir no login Azure CLI global do usuário.
    """

    environment = os.environ.copy()

    AZURE_PROFILE_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    environment["AZURE_CONFIG_DIR"] = str(AZURE_PROFILE_DIR)

    if os.name == "nt":
        environment["AZURE_CORE_ENABLE_BROKER_ON_WINDOWS"] = "true"

    return environment


def run_az(
    arguments: list[str],
    *,
    timeout: int = 900,
) -> dict[str, Any] | list[Any]:
    """
    Executa um comando Azure CLI diretamente pelo Python.
    """

    command = [
        azure_cli(),
        *arguments,
    ]

    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=azure_environment(),
        check=False,
    )

    if completed.returncode != 0:
        detail = (
            completed.stderr
            or completed.stdout
            or "Falha desconhecida do Azure CLI."
        ).strip()

        raise RuntimeError(
            f"Falha no Azure CLI:\n{detail}"
        )

    output = completed.stdout.strip()

    if not output:
        return {}

    try:
        return json.loads(output)

    except json.JSONDecodeError as error:
        raise RuntimeError(
            "O Azure CLI retornou uma resposta inválida.\n\n"
            f"Resposta:\n{output[:2000]}"
        ) from error


# ============================================================================
# SESSÃO AZURE
# ============================================================================

def clear_isolated_session() -> None:
    """
    Limpa somente a sessão Azure CLI criada por este script.

    A sessão Azure CLI global do Windows não é afetada.
    """

    environment = azure_environment()

    subprocess.run(
        [
            azure_cli(),
            "logout",
        ],
        capture_output=True,
        text=True,
        timeout=60,
        env=environment,
        check=False,
    )

    subprocess.run(
        [
            azure_cli(),
            "account",
            "clear",
        ],
        capture_output=True,
        text=True,
        timeout=60,
        env=environment,
        check=False,
    )


def login(
    tenant: str | None = None,
    *,
    reset: bool = True,
) -> None:
    """
    Executa login Microsoft diretamente pelo Python.

    O Azure CLI abrirá automaticamente a autenticação interativa.
    """

    if reset:
        clear_isolated_session()

    print(
        "\n============================================================"
    )
    print(
        " AUTENTICAÇÃO MICROSOFT"
    )
    print(
        "============================================================"
    )

    print(
        "\nAbrindo autenticação Microsoft..."
    )

    arguments = [
        "login",
        "--allow-no-subscriptions",
        "--output",
        "json",
    ]

    if tenant:
        arguments.extend(
            [
                "--tenant",
                tenant,
            ]
        )

    payload = run_az(
        arguments,
        timeout=900,
    )

    if not payload:
        raise RuntimeError(
            "O login Microsoft não retornou uma sessão válida."
        )

    print(
        "\nAutenticação Microsoft concluída."
    )


def current_account() -> dict[str, Any]:
    """
    Obtém informações da conta Azure atualmente autenticada.
    """

    payload = run_az(
        [
            "account",
            "show",
            "--output",
            "json",
        ],
        timeout=60,
    )

    if not isinstance(payload, dict):
        raise RuntimeError(
            "Não foi possível identificar a conta Azure autenticada."
        )

    return payload


def print_current_account() -> None:
    """
    Exibe usuário e tenant utilizados pela análise.
    """

    account = current_account()

    user = account.get("user")

    username = (
        user.get("name")
        if isinstance(user, dict)
        else None
    )

    tenant_id = account.get("tenantId")

    print(
        "\n============================================================"
    )
    print(
        " SESSÃO AZURE"
    )
    print(
        "============================================================"
    )

    print(
        f"Usuário : {username or 'Não identificado'}"
    )

    print(
        f"Tenant  : {tenant_id or 'Não identificado'}"
    )


# ============================================================================
# TOKEN POWER BI
# ============================================================================

def access_token(
    tenant: str | None = None,
) -> str:
    """
    Obtém token delegado para o Power BI REST API.
    """

    arguments = [
        "account",
        "get-access-token",
        "--resource",
        POWER_BI_RESOURCE,
        "--output",
        "json",
    ]

    if tenant:
        arguments.extend(
            [
                "--tenant",
                tenant,
            ]
        )

    payload = run_az(
        arguments,
        timeout=120,
    )

    if not isinstance(payload, dict):
        raise RuntimeError(
            "O Azure CLI retornou uma resposta inesperada "
            "ao solicitar o token Power BI."
        )

    token = payload.get("accessToken")

    if not isinstance(token, str) or not token:
        raise RuntimeError(
            "O Azure CLI não retornou um access token válido "
            "para o Power BI."
        )

    expires_on = (
        payload.get("expiresOn")
        or payload.get("expires_on")
    )

    print(
        "\nToken Power BI obtido com sucesso."
    )

    if expires_on:
        print(
            f"Expiração: {expires_on}"
        )

    return token


# ============================================================================
# POWER BI REST API
# ============================================================================

def api_get(
    url: str,
    token: str,
    tenant: str | None,
    *,
    attempts: int = 6,
) -> tuple[dict[str, Any], str]:
    """
    Executa GET na Power BI REST API.

    Comportamentos:
      - renova token uma vez em HTTP 401;
      - respeita Retry-After em HTTP 429;
      - aplica retry exponencial para erros de rede.
    """

    current_token = token

    renewed = False

    for attempt in range(attempts):

        request = urllib.request.Request(
            url,
            headers={
                "Authorization": f"Bearer {current_token}",
                "Accept": "application/json",
                "User-Agent": "AssessmentAnalyzer-GatewayStatus/1.0",
            },
            method="GET",
        )

        try:

            with urllib.request.urlopen(
                request,
                timeout=120,
            ) as response:

                payload = json.load(response)

                if not isinstance(payload, dict):
                    raise RuntimeError(
                        "A API Power BI retornou JSON em formato inesperado."
                    )

                return payload, current_token

        except urllib.error.HTTPError as error:

            body = error.read().decode(
                "utf-8",
                errors="replace",
            )

            # ------------------------------------------------------------
            # 401 - Token inválido / expirado
            # ------------------------------------------------------------

            if error.code == 401:

                if not renewed:

                    print(
                        "\nHTTP 401 recebido."
                    )

                    print(
                        "Tentando renovar o token Power BI..."
                    )

                    current_token = access_token(
                        tenant
                    )

                    renewed = True

                    continue

                raise RuntimeError(
                    "\nPower BI retornou HTTP 401 Unauthorized "
                    "mesmo após a renovação do token.\n\n"
                    "Verifique:\n"
                    "  - se o usuário autenticado pertence ao tenant correto;\n"
                    "  - se possui acesso ao Gateway Cluster;\n"
                    "  - se o tenant corresponde ao Power BI Service;\n"
                    "  - se a conta possui permissão para consultar "
                    "Gateway Clusters.\n\n"
                    f"Resposta Power BI:\n{body[:2000]}"
                ) from error

            # ------------------------------------------------------------
            # 403
            # ------------------------------------------------------------

            if error.code == 403:

                raise RuntimeError(
                    "\nPower BI retornou HTTP 403 Forbidden.\n\n"
                    "O usuário foi autenticado, porém não possui "
                    "permissão suficiente para consultar esta API.\n\n"
                    f"Resposta Power BI:\n{body[:2000]}"
                ) from error

            # ------------------------------------------------------------
            # 429 - Throttling
            # ------------------------------------------------------------

            if (
                error.code == 429
                and attempt + 1 < attempts
            ):

                retry_after = error.headers.get(
                    "Retry-After",
                    "5",
                )

                try:

                    delay = max(
                        1,
                        min(
                            120,
                            int(float(retry_after)),
                        ),
                    )

                except ValueError:

                    delay = 5

                print(
                    f"\nPower BI aplicou throttling."
                    f" Nova tentativa em {delay}s..."
                )

                time.sleep(delay)

                continue

            raise RuntimeError(
                f"\nPower BI API HTTP {error.code}:\n"
                f"{body[:2000]}"
            ) from error

        except urllib.error.URLError as error:

            if attempt + 1 >= attempts:

                raise RuntimeError(
                    "Falha de rede ao consultar o Power BI:\n"
                    f"{error.reason}"
                ) from error

            delay = min(
                30,
                2 ** attempt,
            )

            print(
                f"\nFalha de rede."
                f" Nova tentativa em {delay}s..."
            )

            time.sleep(delay)

    raise RuntimeError(
        "A API Power BI excedeu o número máximo de tentativas."
    )


# ============================================================================
# TRATAMENTO DOS DADOS
# ============================================================================

def rows(
    payload: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Localiza a coleção principal da resposta.
    """

    for key in (
        "value",
        "items",
        "data",
    ):

        value = payload.get(key)

        if isinstance(value, list):

            return [
                item
                for item in value
                if isinstance(item, dict)
            ]

    return []


def normalized_status(
    value: object,
) -> str:
    """
    Normaliza os diferentes possíveis estados de gateway.
    """

    status = (
        str(value or "")
        .strip()
        .casefold()
        .replace("_", "")
        .replace("-", "")
    )

    if status in {
        "online",
        "healthy",
        "active",
        "connected",
        "running",
        "live",
        "enabled",
        "installed",
    }:
        return "online"

    if status in {
        "offline",
        "unavailable",
        "disconnected",
        "inactive",
        "disabled",
        "failed",
        "notinstalled",
        "uninstalled",
    }:
        return "offline"

    if status in {
        "warning",
        "degraded",
        "attention",
        "limited",
    }:
        return "warning"

    return "unknown"


# ============================================================================
# COLETA DOS GATEWAYS
# ============================================================================

def collect(
    token: str,
    tenant: str | None = None,
) -> dict[str, Any]:
    """
    Consulta todos os Gateway Clusters e gera resumo de status.
    """

    clusters: list[dict[str, Any]] = []

    visited: set[str] = set()

    url: str | None = GATEWAY_STATUS_URL

    while url:

        if url in visited:

            raise RuntimeError(
                "A paginação da API entrou em ciclo."
            )

        parsed = urllib.parse.urlparse(
            url
        )

        if (
            parsed.scheme != "https"
            or parsed.hostname != "api.powerbi.com"
        ):

            raise RuntimeError(
                "A API retornou uma URL de paginação inválida."
            )

        visited.add(
            url
        )

        print(
            "\nConsultando Gateway Clusters..."
        )

        payload, token = api_get(
            url,
            token,
            tenant,
        )

        clusters.extend(
            rows(payload)
        )

        next_url = (
            payload.get("@odata.nextLink")
            or payload.get("nextLink")
            or payload.get("continuationUri")
        )

        url = (
            next_url
            if isinstance(next_url, str)
            and next_url
            else None
        )

    statuses: list[str] = []

    for cluster in clusters:

        members = cluster.get(
            "memberGateways"
        )

        if not isinstance(
            members,
            list,
        ):
            continue

        for member in members:

            if not isinstance(
                member,
                dict,
            ):
                continue

            raw_status = (
                member.get("state")
                or member.get("gatewayStatus")
                or member.get("status")
            )

            statuses.append(
                normalized_status(
                    raw_status
                )
            )

    counts = Counter(
        statuses
    )

    return {
        "clusters": len(clusters),
        "gateways": len(statuses),
        "online": counts["online"],
        "offline": counts["offline"],
        "warning": counts["warning"],
        "unknown": counts["unknown"],
    }


# ============================================================================
# ARGUMENTOS
# ============================================================================

def parse_args() -> argparse.Namespace:
    """
    Argumentos opcionais.

    Para executar apenas pelo botão Run do VS Code,
    nenhum argumento é obrigatório.
    """

    parser = argparse.ArgumentParser(
        description=(
            "Gateway Status via Login Azure"
        )
    )

    parser.add_argument(
        "--tenant",
        help=(
            "Tenant ID ou domínio Microsoft Entra. "
            "Opcional."
        ),
    )

    parser.add_argument(
        "--reuse-session",
        action="store_true",
        help=(
            "Reutiliza a sessão Azure criada anteriormente."
        ),
    )

    parser.add_argument(
        "--json",
        action="store_true",
        help=(
            "Exibe apenas o JSON final."
        ),
    )

    return parser.parse_args()


# ============================================================================
# EXECUÇÃO
# ============================================================================

def main() -> int:

    args = parse_args()

    try:

        print(
            "\n============================================================"
        )
        print(
            " ASSESSMENT ANALYZER - GATEWAY STATUS"
        )
        print(
            "============================================================"
        )

        # ------------------------------------------------------------
        # LOGIN
        # ------------------------------------------------------------

        if not args.reuse_session:

            login(
                args.tenant,
                reset=True,
            )

        else:

            print(
                "\nReutilizando sessão Microsoft existente..."
            )

        # ------------------------------------------------------------
        # CONTA
        # ------------------------------------------------------------

        print_current_account()

        # ------------------------------------------------------------
        # TOKEN
        # ------------------------------------------------------------

        token = access_token(
            args.tenant
        )

        # ------------------------------------------------------------
        # POWER BI
        # ------------------------------------------------------------

        result = collect(
            token,
            args.tenant,
        )

        # ------------------------------------------------------------
        # RESULTADO
        # ------------------------------------------------------------

        print(
            "\n============================================================"
        )
        print(
            " RESULTADO"
        )
        print(
            "============================================================"
        )

        if args.json:

            print(
                json.dumps(
                    result,
                    ensure_ascii=False,
                    indent=2,
                )
            )

        else:

            print(
                f"Clusters:             {result['clusters']}"
            )

            print(
                f"Gateways / Nós:       {result['gateways']}"
            )

            print(
                f"Online:               {result['online']}"
            )

            print(
                f"Offline:              {result['offline']}"
            )

            print(
                f"Atenção:              {result['warning']}"
            )

            print(
                f"Status desconhecido:  {result['unknown']}"
            )

        print(
            "\nConsulta concluída com sucesso."
        )

        return 0

    except KeyboardInterrupt:

        print(
            "\nOperação cancelada pelo usuário.",
            file=sys.stderr,
        )

        return 130

    except Exception as error:

        print(
            "\n============================================================",
            file=sys.stderr,
        )

        print(
            " ERRO",
            file=sys.stderr,
        )

        print(
            "============================================================",
            file=sys.stderr,
        )

        print(
            str(error),
            file=sys.stderr,
        )

        return 1


if __name__ == "__main__":
    raise SystemExit(main())