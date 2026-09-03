from gns_app.services.windows_focus import (
    _is_outlook_certificate_yes_button,
    _looks_like_outlook_certificate_warning,
)


def test_matches_only_known_outlook_certificate_warning_text():
    assert _looks_like_outlook_certificate_warning(
        "Предупреждение системы безопасности в Интернете",
        [
            "Сертификат безопасности сервера просрочен.",
            "Продолжить использование этого сервера?",
        ],
    )
    assert _looks_like_outlook_certificate_warning(
        "Internet Security Warning",
        [
            "The security certificate has expired.",
            "Do you want to continue using this server?",
        ],
    )
    assert _looks_like_outlook_certificate_warning(
        "Предупреждение безопасности Интернета",
        [
            "Сервер использует сертификат безопасности, который не может "
            "быть проверен.",
            "Обработка прервана в корневом сертификате, у которого отсутствует "
            "отношение доверия с поставщиком доверия.",
            "Продолжать использовать этот сервер?",
        ],
    )


def test_does_not_match_unrelated_confirmation_dialogs():
    assert not _looks_like_outlook_certificate_warning(
        "Microsoft Outlook",
        ["Удалить выбранное сообщение?"],
    )
    assert not _looks_like_outlook_certificate_warning(
        "Предупреждение системы безопасности",
        ["Разрешить запуск макроса?"],
    )
    assert not _looks_like_outlook_certificate_warning(
        "Предупреждение системы безопасности",
        ["Сертификат действителен."],
    )


def test_russian_outlook_yes_button_is_recognized_without_standard_id():
    assert _is_outlook_certificate_yes_button("Да", 1008)
    assert _is_outlook_certificate_yes_button("&Yes", 1008)
    assert not _is_outlook_certificate_yes_button("Нет", 1008)
