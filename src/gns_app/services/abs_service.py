from __future__ import annotations

from collections.abc import Callable
from urllib.error import HTTPError, URLError

from gns_app.config import Settings
from gns_app.domain import AbsCheckResult, AbsStatus
from gns_app.vendor.tolubay_hub import (
    AuthenticationError,
    ProtocolError,
    TolubayClient,
    TolubayConfig,
    TolubayError,
)


class FakeAbsGateway:
    """Фейковая АБС.

    Учётные данные проверяются только в текущем вызове и не передаются
    в базу или журнал. Значения `offline` и `invalid` позволяют проверить
    сценарии ошибок.
    """

    FOUND_INNS = {
        "11111111111111",
        "22222222222222",
    }
    is_fake = True
    supports_session = True

    def check(
        self,
        username: str,
        password: str,
        taxpayers: list[dict[str, str]],
    ) -> AbsCheckResult:
        if not username or not password:
            return AbsCheckResult(
                status=AbsStatus.AUTH_ERROR,
                taxpayers=[],
                message="Логин и пароль обязательны.",
            )
        if username.casefold() == "offline":
            return AbsCheckResult(
                status=AbsStatus.UNAVAILABLE,
                taxpayers=[],
                message="Тестовый сценарий: АБС недоступна.",
            )
        if password.casefold() == "invalid":
            return AbsCheckResult(
                status=AbsStatus.AUTH_ERROR,
                taxpayers=[],
                message="Тестовый сценарий: неверные учётные данные.",
            )

        results: list[dict[str, str]] = []
        found_any = False
        for taxpayer in taxpayers:
            inn = taxpayer["inn"]
            found = inn in self.FOUND_INNS
            found_any = found_any or found
            results.append(
                {
                    "inn": inn,
                    "name": taxpayer["name"],
                    "result": "found" if found else "not_found",
                }
            )

        return AbsCheckResult(
            status=AbsStatus.FOUND if found_any else AbsStatus.NOT_FOUND,
            taxpayers=results,
            message=(
                "В тестовой АБС найден хотя бы один налогоплательщик."
                if found_any
                else "В тестовой АБС налогоплательщики не найдены."
            ),
        )


class TolubayAbsGateway:
    """Read-only адаптер поиска клиентов и счетов в Tolubay ABS."""

    is_fake = False
    supports_session = False

    def __init__(
        self,
        config: TolubayConfig,
        *,
        client_factory: Callable[[TolubayConfig], TolubayClient] = TolubayClient,
    ) -> None:
        if not config.verify_tls:
            raise ValueError("Проверку TLS для рабочей АБС отключать нельзя")
        self.config = config
        self._client_factory = client_factory

    @staticmethod
    def _failure(status: AbsStatus, message: str) -> AbsCheckResult:
        return AbsCheckResult(
            status=status,
            taxpayers=[],
            message=message,
            is_fake=False,
        )

    @staticmethod
    def _normalized_inn(value: object) -> str:
        return "".join(character for character in str(value) if character.isdigit())

    @staticmethod
    def _protocol_status(error: ProtocolError) -> AbsStatus:
        cause = error.__cause__
        if isinstance(cause, URLError) and not isinstance(cause, HTTPError):
            return AbsStatus.UNAVAILABLE
        if isinstance(getattr(cause, "reason", None), (TimeoutError, OSError)):
            return AbsStatus.UNAVAILABLE
        return AbsStatus.TECHNICAL_ERROR

    def check(
        self,
        username: str,
        password: str,
        taxpayers: list[dict[str, str]],
    ) -> AbsCheckResult:
        if not username.strip() or not password:
            return self._failure(
                AbsStatus.AUTH_ERROR,
                "Логин и пароль АБС обязательны.",
            )
        if not taxpayers:
            return self._failure(
                AbsStatus.TECHNICAL_ERROR,
                "Не переданы налогоплательщики для проверки АБС.",
            )
        for taxpayer in taxpayers:
            inn = str(taxpayer.get("inn") or "").strip()
            if len(inn) != 14 or not inn.isdigit():
                return self._failure(
                    AbsStatus.TECHNICAL_ERROR,
                    "АБС получила неподтверждённый ИНН. Проверка остановлена.",
                )

        try:
            client = self._client_factory(self.config)
            client.login(username.strip(), password)
            taxpayer_results: list[dict[str, str]] = []
            for taxpayer in taxpayers:
                inn = taxpayer["inn"]
                matches = client.search_customers(
                    {"SearchIdentificationNo": inn},
                    page_size=100,
                )
                exact_matches = [
                    customer
                    for customer in matches
                    if self._normalized_inn(customer.identity) == inn
                ]
                if not matches:
                    taxpayer_result = AbsStatus.NOT_FOUND
                elif len(matches) != 1 or len(exact_matches) != 1:
                    taxpayer_result = AbsStatus.MULTIPLE
                else:
                    # До уточнения бизнес-правила сама точная анкета клиента
                    # достаточна для ручной обработки. Отсутствие счетов не
                    # разрешает автоматический ответ об отсутствии клиента.
                    taxpayer_result = AbsStatus.FOUND
                taxpayer_results.append(
                    {
                        "inn": inn,
                        "name": taxpayer["name"],
                        "result": str(taxpayer_result),
                    }
                )
        except AuthenticationError:
            return self._failure(
                AbsStatus.AUTH_ERROR,
                "АБС отклонила логин или пароль.",
            )
        except ProtocolError as exc:
            status = self._protocol_status(exc)
            return self._failure(
                status,
                (
                    "АБС недоступна. Отсутствие клиента не подтверждено."
                    if status == AbsStatus.UNAVAILABLE
                    else "АБС вернула неизвестный ответ. Проверка остановлена."
                ),
            )
        except (TimeoutError, OSError):
            return self._failure(
                AbsStatus.UNAVAILABLE,
                "АБС недоступна. Отсутствие клиента не подтверждено.",
            )
        except (TolubayError, ValueError, TypeError, KeyError):
            return self._failure(
                AbsStatus.TECHNICAL_ERROR,
                "Проверка АБС завершилась технической ошибкой.",
            )

        result_values = {item["result"] for item in taxpayer_results}
        if str(AbsStatus.MULTIPLE) in result_values:
            status = AbsStatus.MULTIPLE
            message = "АБС вернула неоднозначное совпадение. Нужна ручная проверка."
        elif str(AbsStatus.FOUND) in result_values:
            status = AbsStatus.FOUND
            message = "В АБС найдена анкета клиента. Нужна ручная обработка."
        else:
            status = AbsStatus.NOT_FOUND
            message = "В АБС точные анкеты клиентов не найдены."
        return AbsCheckResult(
            status=status,
            taxpayers=taxpayer_results,
            message=message,
            is_fake=False,
        )


def create_abs_gateway(settings: Settings) -> FakeAbsGateway | TolubayAbsGateway:
    if settings.abs_mode == "fake":
        return FakeAbsGateway()
    if settings.abs_mode == "tolubay":
        return TolubayAbsGateway(
            TolubayConfig(
                base_url=settings.tolubay_base_url,
                ca_file=settings.tolubay_ca_file,
                verify_tls=True,
                timeout_seconds=settings.tolubay_timeout_seconds,
            )
        )
    raise ValueError("Неизвестный режим АБС")
