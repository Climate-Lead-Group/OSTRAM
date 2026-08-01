"""Legacy B1 configuration checks with injectable reporting and stopping."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import Any
import sys


DEMAND_TIMESLICE_MISMATCH = (
    "These variables are differents, so you need check A-O_Demand.xlsx sheet "
    "Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslices"
)
DEMAND_TIMESLICE_DEFINITION_ERROR = (
    "Check the defintion of Timeslices, into A-O_Demand.xlsx sheet Timeslices "
    "and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslice"
)
CAPACITY_TIMESLICE_MISMATCH = (
    "These variables are differents, so you need check A-O_Parametrization.xlsx "
    "sheet Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslices"
)
CAPACITY_TIMESLICE_DEFINITION_ERROR = (
    "Check the defintion of Timeslices, into A-O_Parametrization.xlsx sheet "
    "Timeslices and Config_MOMF_T1_A.yaml variable xtra_scen/Timeslice"
)
YEARSPLIT_TIMESLICE_ERROR = (
    "These variables have inconsistance variable xtra_scen/Timeslice and "
    "xtra_scen/Timeslices"
)
DAYSPLIT_TIME_BRACKET_ERROR = (
    "These variables have inconsistance variable xtra_scen/DailyTimeBracket"
)


def _abort(
    message: str,
    emit: Callable[[str], Any] | None,
    stop: Callable[[], Any] | None,
) -> None:
    if emit is None:
        emit = print
    if stop is None:
        stop = sys.exit
    emit(message)
    stop()


def validate_demand_timeslices(
    configured_timeslices: Sequence[Any],
    timeslice_mode: Any,
    workbook_timeslices: Sequence[Any],
    *,
    emit: Callable[[str], Any] | None = None,
    stop: Callable[[], Any] | None = None,
) -> None:
    """Apply the two predecessor demand-timeslice failure checks."""
    if (
        configured_timeslices != workbook_timeslices
        and timeslice_mode == "Some"
        and workbook_timeslices != []
    ):
        _abort(DEMAND_TIMESLICE_MISMATCH, emit, stop)
        return
    if (
        (timeslice_mode == "Some" and not workbook_timeslices)
        or (timeslice_mode == "All" and workbook_timeslices)
    ):
        _abort(DEMAND_TIMESLICE_DEFINITION_ERROR, emit, stop)


def validate_capacity_timeslices(
    configured_timeslices: Sequence[Any],
    timeslice_mode: Any,
    workbook_timeslices: Sequence[Any],
    *,
    emit: Callable[[str], Any] | None = None,
    stop: Callable[[], Any] | None = None,
) -> None:
    """Apply the two predecessor capacity-factor timeslice checks."""
    if (
        configured_timeslices != workbook_timeslices
        and timeslice_mode == "Some"
        and workbook_timeslices != []
    ):
        _abort(CAPACITY_TIMESLICE_MISMATCH, emit, stop)
        return
    if (
        (timeslice_mode == "Some" and workbook_timeslices == [])
        or (timeslice_mode == "All" and workbook_timeslices != [])
    ):
        _abort(CAPACITY_TIMESLICE_DEFINITION_ERROR, emit, stop)


def validate_yearsplit_timeslices(
    timeslice_mode: Any,
    configured_timeslices: Sequence[Any],
    *,
    emit: Callable[[str], Any] | None = None,
    stop: Callable[[], Any] | None = None,
) -> None:
    if timeslice_mode == "Some" and configured_timeslices == []:
        _abort(YEARSPLIT_TIMESLICE_ERROR, emit, stop)


def validate_daysplit_time_brackets(
    daily_time_brackets: Sequence[Any],
    configured_timeslices: Sequence[Any],
    *,
    emit: Callable[[str], Any] | None = None,
    stop: Callable[[], Any] | None = None,
) -> None:
    """Retain the predecessor check without repairing its one-bracket branch."""
    if len(daily_time_brackets) < 2 and configured_timeslices == []:
        _abort(DAYSPLIT_TIME_BRACKET_ERROR, emit, stop)
