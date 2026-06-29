from __future__ import annotations

from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class Theme:
    name: str
    background: str
    surface: str
    surface_alt: str
    sidebar: str
    sidebar_alt: str
    primary: str
    primary_hover: str
    accent: str
    accent_hover: str
    success: str
    warning: str
    danger: str
    text: str
    text_muted: str
    text_inverse: str
    border: str
    chip: str


ENTERPRISE_DARK = Theme(
    name="Enterprise Dark Blue",
    background="#0B1220",
    surface="#111827",
    surface_alt="#182235",
    sidebar="#07111F",
    sidebar_alt="#0E1B2D",
    primary="#1D4ED8",
    primary_hover="#2563EB",
    accent="#38BDF8",
    accent_hover="#0EA5E9",
    success="#10B981",
    warning="#F59E0B",
    danger="#EF4444",
    text="#F8FAFC",
    text_muted="#A7B0C0",
    text_inverse="#07111F",
    border="#263246",
    chip="#1E293B",
)

PREMIUM_LIGHT = Theme(
    name="Premium Light",
    background="#F4F7FB",
    surface="#FFFFFF",
    surface_alt="#EDF3FA",
    sidebar="#0B2F4F",
    sidebar_alt="#123E66",
    primary="#155C94",
    primary_hover="#0B2F4F",
    accent="#0EA5E9",
    accent_hover="#0284C7",
    success="#0F766E",
    warning="#B7791F",
    danger="#B42318",
    text="#172B4D",
    text_muted="#5B6B82",
    text_inverse="#FFFFFF",
    border="#D6DEE6",
    chip="#EAF4FB",
)

THEMES: Dict[str, Theme] = {
    ENTERPRISE_DARK.name: ENTERPRISE_DARK,
    PREMIUM_LIGHT.name: PREMIUM_LIGHT,
}


def get_theme(name: str | None = None) -> Theme:
    if name and name in THEMES:
        return THEMES[name]
    return ENTERPRISE_DARK
