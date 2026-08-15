from __future__ import annotations

from gns_app.services.taxpayer_service import TaxpayerKind, classify_taxpayer


def test_company_inn_prefixes_are_legal_entities():
    assert classify_taxpayer("", "01412201510188") == TaxpayerKind.LEGAL_ENTITY
    assert classify_taxpayer("", "31412201510188") == TaxpayerKind.LEGAL_ENTITY
    assert classify_taxpayer("", "51412201510188") == TaxpayerKind.LEGAL_ENTITY


def test_personal_pin_prefixes_are_not_sent_to_company_registry():
    assert classify_taxpayer("Иванов Иван", "11412201510188") == TaxpayerKind.INDIVIDUAL
    assert classify_taxpayer("Иванова Ирина", "21412201510188") == TaxpayerKind.INDIVIDUAL


def test_prefix_four_stays_unknown_without_organization_form():
    assert classify_taxpayer("Иностранный гражданин", "41412201510188") == TaxpayerKind.UNKNOWN
    assert classify_taxpayer(
        'Филиал иностранной компании "Альфа"', "41412201510188"
    ) == TaxpayerKind.LEGAL_ENTITY


def test_legal_form_and_personal_pin_are_a_conflict():
    assert classify_taxpayer('ОсОО "Альфа"', "11412201510188") == TaxpayerKind.UNKNOWN
