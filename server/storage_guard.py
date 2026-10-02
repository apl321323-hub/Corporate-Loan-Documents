"""Server-owned write-scope validation for persistent application documents."""

from __future__ import annotations


class StorageScopeError(ValueError):
    """Raised before persistence when a write plan violates its menu scope."""


COMMON_UPLOAD_KEYS = frozenset({
    "company_info", "bs_is", "cashflow", "asset_quality",
    "business_status", "loan_asset", "borrowing",
})


def validated_write_plan(values: dict, allowed_keys: frozenset[str]) -> dict:
    if not isinstance(values, dict):
        raise StorageScopeError("parser result must be an object")
    unexpected = sorted(set(values) - set(allowed_keys))
    if unexpected:
        raise StorageScopeError("write scope contains unexpected keys: " + ", ".join(unexpected))
    errors = sorted(key for key, value in values.items()
                    if isinstance(value, dict) and value.get("error"))
    if errors:
        raise StorageScopeError("parser errors must not be persisted: " + ", ".join(errors))
    if not values:
        raise StorageScopeError("upload contains no supported data")
    return dict(values)
