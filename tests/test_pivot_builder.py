"""
Unit tests for PivotBuilder behavior.
"""

from __future__ import annotations

from typing import List

from pivot_util.constants import DestinationHandling
from pivot_util.fields import DataField
from pivot_util.pivot_builder import PivotBuilder


class DummyWorkbook:
    def __init__(self) -> None:
        self.sheets: List[object] = []
        self.api = object()


def test_generate_pivot_calls_core(monkeypatch) -> None:
    called = {}

    def _fake_generate_pivot(spec) -> None:
        called["spec"] = spec

    monkeypatch.setattr("pivot_util.pivot_builder.generate_pivot", _fake_generate_pivot)

    builder = PivotBuilder(
        workbook=DummyWorkbook(),
        table_name="Table1",
        destination_handling=DestinationHandling.FIND_OR_CREATE,
        data_fields=[DataField(name="Qty", function="sum")],  # type: ignore[arg-type]
    )

    builder.generate_pivot()
    assert called["spec"] is builder
