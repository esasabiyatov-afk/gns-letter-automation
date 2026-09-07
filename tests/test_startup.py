import asyncio
from pathlib import Path
from types import SimpleNamespace


def test_lifespan_yields_before_deferred_maintenance(monkeypatch):
    from gns_app import main

    events: list[str] = []

    class FakeDatabase:
        @staticmethod
        def initialize():
            events.append("database")

    class FakeWorkflow:
        @staticmethod
        def recover_interrupted_abs_checks():
            return 0

        @staticmethod
        def initialize_employee_profiles():
            events.append("employees")

        @staticmethod
        def initialize_gns_offices():
            events.append("offices")

        @staticmethod
        def initialize_gns_office_emails():
            events.append("emails")

        @staticmethod
        def interrupted_upload_ids():
            return []

    async def deferred(_resume_ids):
        events.append("maintenance")

    async def idle_loop():
        events.append("loop")

    monkeypatch.setattr(
        main,
        "settings",
        SimpleNamespace(
            abs_mode="fake",
            runtime_dir=Path("runtime"),
            ensure_directories=lambda: events.append("directories"),
        ),
    )
    monkeypatch.setattr(main, "db", FakeDatabase())
    monkeypatch.setattr(main, "workflow", FakeWorkflow())
    monkeypatch.setattr(main, "record_event", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        main,
        "_finish_startup_in_background",
        deferred,
    )
    monkeypatch.setattr(main, "_abs_automation_loop", idle_loop)
    monkeypatch.setattr(main, "_outlook_automation_loop", idle_loop)
    monkeypatch.setattr(main, "_outlook_sent_status_loop", idle_loop)

    async def exercise():
        async with main.lifespan(None):
            assert events == [
                "directories",
                "database",
                "employees",
                "offices",
                "emails",
            ]
            await asyncio.sleep(0)
            assert "maintenance" in events

    asyncio.run(exercise())
