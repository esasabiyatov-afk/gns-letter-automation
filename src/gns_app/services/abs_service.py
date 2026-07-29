from __future__ import annotations

from gns_app.domain import AbsCheckResult, AbsStatus


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
