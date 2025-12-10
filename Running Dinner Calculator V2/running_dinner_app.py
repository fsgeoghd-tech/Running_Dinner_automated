"""
Running Dinner Planner
======================

This module converts the original Jupyter notebook into a single Python script that
contains all data-processing steps and also provides a simple UI for organisers with
limited technical experience.

Core features:
- reads the team list (`data.xlsx`) and afterparty information (`Afterparty.xlsx`)
- geocodes all addresses via the OpenRouteService API (ORS)
- builds course assignments (starter / main / dessert) and all encounter pairings
- prepares overview and personalised mail texts for every team
- offers both a Streamlit UI and a CLI workflow

Usage hints:
- Install the required packages once: `pip install pandas openrouteservice streamlit`
- Store your ORS API key in the environment variable `ORS_API_KEY` or pass it via UI
- Start the UI with: `streamlit run running_dinner_app.py`
- Or run the CLI mode: `python running_dinner_app.py --config config.json`

The script keeps a local JSON cache (`geocode_cache.json`) so that previously geocoded
addresses do not consume additional API quota. Delete the cache if an address changes.
"""

from __future__ import annotations

import argparse
import json
import os
import textwrap
import time
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
import math

import pandas as pd

try:
    import pydeck as pdk

    PYDECK_AVAILABLE = True
except Exception:  # pragma: no cover - optional dependency
    PYDECK_AVAILABLE = False

try:
    import plotly.graph_objects as go

    PLOTLY_AVAILABLE = True
except Exception:  # pragma: no cover - optional dependency
    go = None  # type: ignore[assignment]
    PLOTLY_AVAILABLE = False

try:
    from openrouteservice import Client as ORSClient
    from openrouteservice.exceptions import ApiError as ORSApiError

    ORS_AVAILABLE = True
except Exception:  # pragma: no cover - we only reach this branch if ORS is missing
    ORSClient = None  # type: ignore[assignment]
    ORSApiError = Exception  # type: ignore[assignment]
    ORS_AVAILABLE = False

try:
    import streamlit as st
    from streamlit.runtime.scriptrunner import get_script_run_ctx

    STREAMLIT_AVAILABLE = True
except Exception:  # pragma: no cover - Streamlit is optional
    STREAMLIT_AVAILABLE = False

# --------------------------------------------------------------------------------------
# Global configuration values
# --------------------------------------------------------------------------------------

MANDATORY_TEAM_COLUMNS: Tuple[str, ...] = (
    "Team Nr.",
    "Name 1",
    "Name 2",
    "Adress",
    "Ring at",
    "will to double",
    "readiness starter",
    "vegan",
    "routeplus_vegan",
    "Allergies or else",
)

MANDATORY_AFTERPARTY_COLUMNS: Tuple[str, ...] = ("Adress",)

SEES_COLUMNS: Tuple[str, ...] = tuple(f"Sees {idx}" for idx in range(1, 7))

DEFAULT_PROFILE = "cycling-regular"
DEFAULT_OUTPUT_DIR = Path("output")
GEOCODE_CACHE_PATH = Path("geocode_cache.json")

ROLE_BASE_COLORS: Dict[str, Tuple[int, int, int]] = {
    "Starter": (66, 133, 244),  # blau
    "Main": (0, 200, 83),  # grün
    "Dessert": (255, 138, 101),  # koralle
}

STAGE_COLORS: Dict[str, Tuple[int, int, int]] = {
    "Zur Vorspeise": (255, 171, 0),  # orange
    "Zur Hauptspeise": (0, 188, 212),  # türkis
    "Zum Dessert": (76, 175, 80),  # dunkelgrün
}


# --------------------------------------------------------------------------------------
# Data transfer objects
# --------------------------------------------------------------------------------------

@dataclass
class PlannerConfig:
    """Container for all organiser inputs that influence the mail texts."""

    language: str = "de"  # "de" or "en"
    time_starter: str = "18:00"
    time_main: str = "20:00"
    time_dessert: str = "22:00"
    time_afterparty: str = "23:30"
    drinks_option: str = "paid"  # "paid" (Spendenbasis) or "free"
    organizer_name: str = ""
    organizer_phone: str = ""
    awareness_enabled: bool = False
    awareness_phone: Optional[str] = None


@dataclass
class PlannerResult:
    """Collection of all artefacts that we produce during planning."""

    address_table: pd.DataFrame
    starters_table: pd.DataFrame
    main_courses_table: pd.DataFrame
    desserts_table: pd.DataFrame
    distance_frames: Dict[str, pd.DataFrame]
    overview_text: str
    mail_contents: Dict[int, str]
    double_cook_teams: List[int]


# --------------------------------------------------------------------------------------
# Utility helpers (I/O, validation, caching)
# --------------------------------------------------------------------------------------

def _load_geocode_cache(cache_path: Path) -> Dict[str, Tuple[float, float]]:
    """Read cached geocoding results so we do not hit the API twice for the same address."""

    if cache_path.exists():
        with cache_path.open("r", encoding="utf-8") as handle:
            raw_cache = json.load(handle)
        return {key: tuple(value) for key, value in raw_cache.items()}
    return {}


def _save_geocode_cache(cache_path: Path, cache: Dict[str, Tuple[float, float]]) -> None:
    """Persist the cache after new addresses were geocoded."""

    with cache_path.open("w", encoding="utf-8") as handle:
        json.dump(cache, handle, indent=2)


def _validate_columns(frame: pd.DataFrame, required: Iterable[str], name: str) -> None:
    """Ensure that all required columns exist before we start processing."""

    missing = [column for column in required if column not in frame.columns]
    if missing:
        raise ValueError(f"{name} fehlt folgende Spalten: {', '.join(missing)}")


def _read_excel(source: BytesIO | str | Path) -> pd.DataFrame:
    """Read an Excel file from disk or an uploaded stream."""

    return pd.read_excel(source)


def _get_ors_client(api_key: str) -> ORSClient:
    """Create an ORS client with clear error reporting."""

    if not ORS_AVAILABLE:
        raise ImportError(
            "Das Paket 'openrouteservice' ist nicht installiert. "
            "Installiere es mit 'pip install openrouteservice'."
        )
    if not api_key:
        raise ValueError(
            "Es wurde kein OpenRouteService-API-Key gefunden. "
            "Hinterlege ihn im UI oder setze die Umgebungsvariable ORS_API_KEY."
        )
    return ORSClient(key=api_key)


# --------------------------------------------------------------------------------------
# Geocoding logic
# --------------------------------------------------------------------------------------

def _geocode_address(
    address: str,
    client: ORSClient,
    cache: Dict[str, Tuple[float, float]],
    pause_seconds: float,
) -> Tuple[float, float]:
    """
    Geocode a single address (with caching and a gentle delay to respect rate limits).

    Parameters
    ----------
    address: str
        The address that should be geocoded.
    client: ORSClient
        OpenRouteService client instance.
    cache: Dict[str, Tuple[float, float]]
        Mutable cache of already resolved addresses.
    pause_seconds: float
        Sleep duration between API calls (helps with rate limits).
    """

    if address in cache:
        return cache[address]

    # Waiting before the request avoids hitting the free-tier rate limit.
    time.sleep(pause_seconds)
    try:
        response = client.pelias_search(text=address)
        coordinates = response["features"][0]["geometry"]["coordinates"]
        lon, lat = float(coordinates[0]), float(coordinates[1])
    except (KeyError, IndexError, TypeError, ORSApiError) as exc:  # noqa: PERF203 - explicit exceptions help debugging
        raise RuntimeError(f"Adresse '{address}' konnte nicht geokodiert werden.") from exc

    cache[address] = (lon, lat)
    return lon, lat


def _geocode_dataframe(
    frame: pd.DataFrame,
    client: ORSClient,
    cache: Dict[str, Tuple[float, float]],
    pause_seconds: float,
) -> pd.DataFrame:
    """Attach longitude/latitude columns to a dataframe with an 'Adress' column."""

    result = frame.copy()
    longitudes: List[float] = []
    latitudes: List[float] = []

    for address in result["Adress"]:
        lon, lat = _geocode_address(str(address), client, cache, pause_seconds)
        longitudes.append(lon)
        latitudes.append(lat)

    result["longitude"] = longitudes
    result["latitude"] = latitudes
    return result


# --------------------------------------------------------------------------------------
# Distance and assignment helpers
# --------------------------------------------------------------------------------------

def _distance_matrix(
    client: ORSClient,
    origins: List[List[float]],
    destinations: List[List[float]],
    profile: str = DEFAULT_PROFILE,
) -> Tuple[List[List[float]], List[List[float]]]:
    """Wrapper around the ORS distance_matrix endpoint with friendly error messages."""

    try:
        matrix = client.distance_matrix(
            locations=origins + destinations,
            profile=profile,
            sources=list(range(len(origins))),
            destinations=list(range(len(origins), len(origins) + len(destinations))),
            metrics=["distance", "duration"],
        )
    except ORSApiError as exc:
        raise RuntimeError("OpenRouteService Distance-Matrix Anfrage fehlgeschlagen.") from exc

    return matrix["distances"], matrix["durations"]


def _prepare_distance_to_afterparty(
    address_table: pd.DataFrame,
    afterparty_table: pd.DataFrame,
    client: ORSClient,
) -> pd.DataFrame:
    """Compute biking distances from each team to the afterparty."""

    team_coords = address_table[["longitude", "latitude"]].to_numpy().tolist()
    afterparty_coord = afterparty_table[["longitude", "latitude"]].iloc[0].tolist()
    distances, durations = _distance_matrix(client, team_coords, [afterparty_coord])

    distance_df = address_table[
        [
            "Team Nr.",
            "Adress",
            "will to double",
            "readiness starter",
            "vegan",
            "routeplus_vegan",
            "latitude",
            "longitude",
        ]
    ].copy()
    distance_df["Distance_km"] = [segment[0] / 1000 for segment in distances]
    distance_df["Duration_H"] = [segment[0] / 3600 for segment in durations]
    distance_df.sort_values("Duration_H", ascending=True, inplace=True)
    distance_df.reset_index(drop=True, inplace=True)
    return distance_df


def _select_double_cooks(distance_df: pd.DataFrame, count: int) -> List[int]:
    """Pick the teams with the highest willingness to double-cook."""

    if count <= 0:
        return []

    candidates = (
        distance_df.sort_values(
            by=["will to double", "Duration_H"],
            ascending=[False, True],
        )["Team Nr."]
        .astype(int)
        .tolist()
    )
    # Deduplicate while preserving order.
    seen: set[int] = set()
    unique_candidates: List[int] = []
    for team_nr in candidates:
        if team_nr not in seen:
            unique_candidates.append(team_nr)
            seen.add(team_nr)
        if len(unique_candidates) == count:
            break

    return unique_candidates[:count]


def _assign_courses(
    distance_df: pd.DataFrame,
    double_cook_numbers: List[int],
    double_cook_quota: int,
) -> Tuple[List[int], List[int], List[int]]:
    """
    Create the starter / main / dessert host lists.

    The logic mirrors the original notebook but is packaged in a reusable function.
    Dessert hosts are chosen closest to the afterparty, starter hosts among teams that
    indicated readiness, and remaining teams are distributed to the main course.
    """

    desserts: List[int] = []
    starters: List[int] = []
    main_courses: List[int] = []
    double_courses: List[int] = []

    team_count = len(distance_df)
    double_cook_quota = max(0, min(double_cook_quota, 2))  # realistically max two extras
    course_width = team_count // 3 + (1 if double_cook_quota else 0)

    # Dessert hosts: best positioned towards the afterparty
    for _, row in distance_df.head(course_width).iterrows():
        team_nr = int(row["Team Nr."])
        desserts.append(team_nr)
        if team_nr in double_cook_numbers and len(double_courses) < double_cook_quota:
            double_courses.append(team_nr)
            main_courses.append(team_nr)

    # Starter hosts: only teams with explicit readiness, prioritising those farther away
    starters_needed = course_width - len([team for team in double_courses if team in starters])
    dessert_prep = distance_df[distance_df["readiness starter"] == True].copy()  # noqa: E712 - explicit boolean comparison
    for _, row in dessert_prep[::-1].iterrows():
        team_nr = int(row["Team Nr."])
        if team_nr in desserts:
            continue
        if len(starters) >= starters_needed:
            break
        starters.append(team_nr)
        if team_nr in double_cook_numbers and len(double_courses) < double_cook_quota:
            double_courses.append(team_nr)
            main_courses.append(team_nr)

    # Main course hosts: everyone who is still unassigned
    already_assigned = set(starters + desserts + main_courses)
    for _, row in distance_df.iterrows():
        team_nr = int(row["Team Nr."])
        if team_nr in already_assigned:
            continue
        main_courses.append(team_nr)
        already_assigned.add(team_nr)
        if team_nr in double_cook_numbers and len(double_courses) < double_cook_quota:
            double_courses.append(team_nr)
            starters.append(team_nr)
            already_assigned.add(team_nr)

    # Safety: ensure we have at least one host in each category
    if not starters or not desserts or not main_courses:
        raise RuntimeError("Die Kursaufteilung ist fehlgeschlagen. Prüfe die Eingabedaten.")

    return starters, main_courses, desserts


def _build_course_table(distance_df: pd.DataFrame, team_numbers: List[int]) -> pd.DataFrame:
    """Return a table that preserves the original order of selected teams."""

    ordered_index = pd.Index(team_numbers, name="Team Nr.")
    frame = (
        distance_df.set_index("Team Nr.")
        .loc[ordered_index]
        .reset_index()
        .copy()
    )
    return frame


def _many_to_many_distances(
    client: ORSClient,
    origin_table: pd.DataFrame,
    destination_table: pd.DataFrame,
    origin_label: str,
    destination_label: str,
) -> pd.DataFrame:
    """Compute the distance matrix between two sets of teams."""

    origin_coords = origin_table[["longitude", "latitude"]].to_numpy().tolist()
    destination_coords = destination_table[["longitude", "latitude"]].to_numpy().tolist()

    if not origin_coords or not destination_coords:
        raise RuntimeError("Für die Distanzmatrix fehlen Ursprungs- oder Zielteams.")

    distances, durations = _distance_matrix(client, origin_coords, destination_coords)

    rows: List[Dict[str, float | int]] = []
    for origin_idx, origin_row in origin_table.iterrows():
        for destination_idx, destination_row in destination_table.iterrows():
            rows.append(
                {
                    origin_label: int(origin_row["Team Nr."]),
                    destination_label: int(destination_row["Team Nr."]),
                    "Distance_km": distances[origin_idx][destination_idx] / 1000,
                    "Duration_H": durations[origin_idx][destination_idx] / 3600,
                }
            )

    return pd.DataFrame(rows)


def _assignment(
    address_table: pd.DataFrame,
    distance_afterparty: pd.DataFrame,
    distance_between_groups: pd.DataFrame,
    prioritize_center: bool,
    sees_first: str,
    sees_second: str,
    start_column: str,
    target_column: str,
) -> None:
    """
    Populate the encounter columns (`Sees 1` ... `Sees 6`).

    This is mostly a direct port from the notebook with additional safeguards against
    chained assignment warnings.
    """

    sorted_afterparty = distance_afterparty.sort_values(
        "Duration_H",
        ascending=prioritize_center,
    )

    for _, team_row in sorted_afterparty.iterrows():
        team_id = int(team_row["Team Nr."])
        team_info = address_table.loc[team_id]

        if team_info[sees_first] != 0:
            # Already assigned, skip
            continue

        potential_guests = distance_between_groups.loc[
            distance_between_groups[start_column] == team_id
        ]

        # Skip if this team would host itself in the next step
        if potential_guests[start_column].isin(potential_guests[target_column]).any():
            continue

        nogos = {int(value) for value in team_info[list(SEES_COLUMNS)] if int(value) != 0}

        relevant_targets = potential_guests.sort_values("Duration_H", ascending=True)
        for _, target_row in relevant_targets.iterrows():
            target_id = int(target_row[target_column])
            if target_id == team_id or target_id in nogos:
                continue

            target_info = address_table.loc[target_id]
            if target_info[sees_first] == 0:
                address_table.loc[target_id, sees_first] = team_id
                address_table.loc[team_id, sees_first] = target_id
                break

            if target_info[sees_second] != 0:
                continue

            third_team = int(target_info[sees_first])
            if third_team in nogos or third_team == 0:
                continue

            address_table.loc[target_id, sees_second] = team_id
            address_table.loc[team_id, sees_first] = target_id
            address_table.loc[team_id, sees_second] = third_team
            address_table.loc[third_team, sees_second] = team_id
            break


def _overwrite_routes(
    guest_distances: pd.DataFrame,
    host_distances: pd.DataFrame,
    address_table: pd.DataFrame,
    guest_start: str,
    guest_target: str,
    host_start: str,
    host_target: str,
    sees_first: str,
) -> pd.DataFrame:
    """
    Copy the main-host route information onto their guests so everyone shares the same
    travel duration for the next leg.
    """

    result = guest_distances.copy()
    for guest_team, guest_rows in result.groupby(guest_start):
        guest_team = int(guest_team)
        if address_table.loc[guest_team, sees_first] != int(
            guest_rows.iloc[0][guest_target]
        ):
            continue

        relevant_hosts = host_distances.loc[
            host_distances[host_start] == int(guest_rows.iloc[0][guest_target])
        ]

        for (guest_idx, _), (host_idx, _) in zip(guest_rows.iterrows(), relevant_hosts.iterrows()):
            result.loc[guest_idx, guest_target] = int(host_distances.loc[host_idx, host_target])
            result.loc[guest_idx, "Duration_H"] = float(host_distances.loc[host_idx, "Duration_H"])

    return result


# --------------------------------------------------------------------------------------
# Reporting helpers
# --------------------------------------------------------------------------------------

def _build_overview_text(
    address_table: pd.DataFrame,
    starters: List[int],
    main_courses: List[int],
    desserts: List[int],
) -> str:
    """Create the overview.txt content."""

    def _sentence(team_id: int, role: str, first: str, second: str) -> str:
        sees_first = int(address_table.loc[team_id, first])
        sees_second = int(address_table.loc[team_id, second])
        return f"Team Nr.{team_id} ({role}) hosts the Teams {sees_first} and {sees_second}."

    sections: List[str] = ["Starters\n\n"]
    seen_sentences: set[str] = set()
    for team_id in starters:
        sentence = _sentence(team_id, "Starter", "Sees 1", "Sees 2")
        if sentence not in seen_sentences:
            sections.append(f"{sentence}\n\n")
            seen_sentences.add(sentence)

    sections.append("Main Courses\n\n")
    for team_id in main_courses:
        sentence = _sentence(team_id, "Main Course", "Sees 3", "Sees 4")
        if sentence not in seen_sentences:
            sections.append(f"{sentence}\n\n")
            seen_sentences.add(sentence)

    sections.append("Desserts\n\n")
    for team_id in desserts:
        sentence = _sentence(team_id, "Dessert", "Sees 5", "Sees 6")
        if sentence not in seen_sentences:
            sections.append(f"{sentence}\n\n")
            seen_sentences.add(sentence)

    return "".join(sections)


def _format_address_line(row: pd.Series) -> str:
    """Compact helper to format address and doorbell hints."""

    return f"{row['Adress']} (klingeln bei {row['Ring at']})"


def _mail_body_de(
    address_table: pd.DataFrame,
    team_id: int,
    config: PlannerConfig,
    starters: List[int],
    main_courses: List[int],
    desserts: List[int],
    afterparty_row: pd.Series,
) -> str:
    """Generate the German mail body."""

    team = address_table.loc[team_id]
    paragraphs: List[str] = []

    paragraphs.append(
        f"Hallo {team['Name 1']} und {team['Name 2']},\n\n"
        "vielen Dank für eure Anmeldung! Hier ist euer persönlicher Fahrplan für den Abend."
    )

    def _guest_info(sees_column: str) -> pd.Series:
        target_id = int(team[sees_column])
        if target_id == 0:
            return team
        return address_table.loc[target_id]

    def _allergy_text(guest_row: pd.Series) -> str:
        text = str(guest_row.get("Allergies or else", "")).strip()
        return text if text else "keine besonderen Hinweise"

    # Vorspeise Abschnitt
    if team_id in starters:
        guest_1 = _guest_info("Sees 1")
        guest_2 = _guest_info("Sees 2")
        paragraphs.append(
            "Ihr startet als Gastgeber*innen der Vorspeise. "
            f"Bitte beachtet folgende Essgewohnheiten: { _allergy_text(guest_1) } "
            f"und { _allergy_text(guest_2) }."
        )
        paragraphs.append(
            "Plant eure Vorspeise so, dass sie entweder gut vorbereitet werden kann "
            "oder sich schnell finalisieren lässt, sobald eure Gäste eintreffen."
        )
    else:
        host = _guest_info("Sees 1")
        paragraphs.append(
            "Den Abend beginnt ihr als Gäste bei "
            f"{host['Name 1']} und {host['Name 2']} ({_format_address_line(host)})."
        )

    # Hauptspeise Abschnitt
    if team_id in main_courses:
        guest_1 = _guest_info("Sees 3")
        guest_2 = _guest_info("Sees 4")
        paragraphs.append(
            "Zur Hauptspeise seid ihr Gastgeber*innen. "
            f"Essgewohnheiten: { _allergy_text(guest_1) } und { _allergy_text(guest_2) }."
        )
    else:
        host = _guest_info("Sees 3")
        paragraphs.append(
            "Die Hauptspeise genießt ihr bei "
            f"{host['Name 1']} und {host['Name 2']} ({_format_address_line(host)})."
        )

    # Dessert Abschnitt
    if team_id in desserts:
        guest_1 = _guest_info("Sees 5")
        guest_2 = _guest_info("Sees 6")
        paragraphs.append(
            "Zum Abschluss serviert ihr die Nachspeise. "
            f"Essgewohnheiten: { _allergy_text(guest_1) } und { _allergy_text(guest_2) }."
        )
    else:
        host = _guest_info("Sees 5")
        paragraphs.append(
            "Zum Dessert geht es zu "
            f"{host['Name 1']} und {host['Name 2']} ({_format_address_line(host)})."
        )

    schedule = textwrap.dedent(
        f"""
        Zeitplan für euch:
        {config.time_starter} – Vorspeise
        {config.time_main} – Hauptspeise
        {config.time_dessert} – Nachspeise
        """
    ).strip()
    paragraphs.append(schedule)

    if config.drinks_option == "paid":
        drinks_text = (
            f"Ab ca. {config.time_afterparty} Uhr treffen wir uns in "
            f"{afterparty_row['Adress']} zum Ausklang. Getränke gibt es auf Spendenbasis."
        )
    else:
        drinks_text = (
            f"Ab ca. {config.time_afterparty} Uhr treffen wir uns in "
            f"{afterparty_row['Adress']} zum Ausklang. Getränke stellen wir euch kostenlos."
        )
    paragraphs.append(drinks_text)

    closing_lines = [
        "Bitte gebt eurem Kochpartner bzw. eurer Kochpartnerin Bescheid, "
        "da diese Mail nur eine Person erreicht.",
        f"Bei kurzfristigen Änderungen erreicht ihr {config.organizer_name} "
        f"unter {config.organizer_phone}.",
    ]
    if config.awareness_enabled and config.awareness_phone:
        closing_lines.append(f"Das Awareness-Team erreicht ihr unter {config.awareness_phone}.")
    closing_lines.append(
        "Wenn ihr krank werdet und keinen Ersatz findet, meldet euch bitte so früh wie möglich."
    )
    paragraphs.append("\n".join(closing_lines))

    return "\n\n".join(paragraphs).strip() + "\n"


def _mail_body_en(
    address_table: pd.DataFrame,
    team_id: int,
    config: PlannerConfig,
    starters: List[int],
    main_courses: List[int],
    desserts: List[int],
    afterparty_row: pd.Series,
) -> str:
    """Generate the English mail body."""

    team = address_table.loc[team_id]
    paragraphs: List[str] = []

    paragraphs.append(
        f"Hi {team['Name 1']} and {team['Name 2']},\n\n"
        "thanks for joining our running dinner! Here is your personal schedule."
    )

    def _guest_info(sees_column: str) -> pd.Series:
        target_id = int(team[sees_column])
        if target_id == 0:
            return team
        return address_table.loc[target_id]

    def _allergy_text(guest_row: pd.Series) -> str:
        text = str(guest_row.get("Allergies or else", "")).strip()
        return text if text else "no special requirements"

    if team_id in starters:
        guest_1 = _guest_info("Sees 1")
        guest_2 = _guest_info("Sees 2")
        paragraphs.append(
            "You kick off the night by hosting the starter. "
            f"Please keep these dietary notes in mind: { _allergy_text(guest_1) } "
            f"and { _allergy_text(guest_2) }."
        )
        paragraphs.append(
            "Prepare the starter so you can welcome your guests relaxed when they arrive."
        )
    else:
        host = _guest_info("Sees 1")
        paragraphs.append(
            "You start as guests at "
            f"{host['Name 1']} and {host['Name 2']} ({_format_address_line(host)})."
        )

    if team_id in main_courses:
        guest_1 = _guest_info("Sees 3")
        guest_2 = _guest_info("Sees 4")
        paragraphs.append(
            "Next you host the main course. "
            f"Dietary notes: { _allergy_text(guest_1) } and { _allergy_text(guest_2) }."
        )
    else:
        host = _guest_info("Sees 3")
        paragraphs.append(
            "The main course is served by "
            f"{host['Name 1']} and {host['Name 2']} ({_format_address_line(host)})."
        )

    if team_id in desserts:
        guest_1 = _guest_info("Sees 5")
        guest_2 = _guest_info("Sees 6")
        paragraphs.append(
            "Finally you serve dessert. "
            f"Keep in mind: { _allergy_text(guest_1) } and { _allergy_text(guest_2) }."
        )
    else:
        host = _guest_info("Sees 5")
        paragraphs.append(
            "Dessert awaits you at "
            f"{host['Name 1']} and {host['Name 2']} ({_format_address_line(host)})."
        )

    schedule = textwrap.dedent(
        f"""
        Your timeline:
        {config.time_starter} – Starter
        {config.time_main} – Main course
        {config.time_dessert} – Dessert
        """
    ).strip()
    paragraphs.append(schedule)

    if config.drinks_option == "paid":
        drinks_text = (
            f"From about {config.time_afterparty} we meet at {afterparty_row['Adress']} "
            "to wrap up the evening. Drinks are available on a donation basis."
        )
    else:
        drinks_text = (
            f"From about {config.time_afterparty} we meet at {afterparty_row['Adress']} "
            "to wrap up the evening. Drinks are on us."
        )
    paragraphs.append(drinks_text)

    closing_lines = [
        "Please inform your cooking partner – only one person receives this mail.",
        f"For short-notice updates call {config.organizer_name} at {config.organizer_phone}.",
    ]
    if config.awareness_enabled and config.awareness_phone:
        closing_lines.append(f"Our awareness team is available via {config.awareness_phone}.")
    closing_lines.append("If you fall ill and cannot find a replacement, let us know as soon as possible.")
    paragraphs.append("\n".join(closing_lines))

    return "\n\n".join(paragraphs).strip() + "\n"


def _mail_body(
    address_table: pd.DataFrame,
    team_id: int,
    config: PlannerConfig,
    starters: List[int],
    main_courses: List[int],
    desserts: List[int],
    afterparty_row: pd.Series,
) -> str:
    """Dispatch to the correct language implementation."""

    language = (config.language or "de").lower()
    if language.startswith("en"):
        return _mail_body_en(address_table, team_id, config, starters, main_courses, desserts, afterparty_row)
    return _mail_body_de(address_table, team_id, config, starters, main_courses, desserts, afterparty_row)


def _rgb_to_hex(color: Iterable[int]) -> str:
    """Convert RGB tuples/lists (0-255) to hexadecimal strings."""

    r, g, b = [max(0, min(255, int(value))) for value in list(color)[:3]]
    return f"#{r:02x}{g:02x}{b:02x}"


def _blend_role_color(roles: List[str]) -> List[int]:
    """Average the base colors of all roles a team fulfils."""

    filtered_roles = [role for role in roles if role in ROLE_BASE_COLORS]
    if not filtered_roles:
        return [180, 180, 180]

    summed = [0, 0, 0]
    for role in filtered_roles:
        base_color = ROLE_BASE_COLORS[role]
        for idx in range(3):
            summed[idx] += base_color[idx]

    count = len(filtered_roles)
    return [int(summed[idx] / count) for idx in range(3)]


def prepare_map_frames(
    address_table: pd.DataFrame,
    starters: Iterable[int],
    main_courses: Iterable[int],
    desserts: Iterable[int],
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Build the point and edge dataframes used for the map visualisation.

    Points represent team home bases; edges represent the routes teams travel to their
    hosts at each course.
    """

    starters_set = {int(value) for value in starters}
    main_set = {int(value) for value in main_courses}
    dessert_set = {int(value) for value in desserts}

    table = address_table.copy()
    table["Team Nr."] = table["Team Nr."].astype(int)

    def _roles_for_team(team_id: int) -> List[str]:
        roles: List[str] = []
        if team_id in starters_set:
            roles.append("Starter")
        if team_id in main_set:
            roles.append("Main")
        if team_id in dessert_set:
            roles.append("Dessert")
        return roles

    table["roles_list"] = table["Team Nr."].apply(_roles_for_team)
    table["role"] = table["roles_list"].apply(lambda roles: " & ".join(roles) if roles else "Gast")
    table["color"] = table["roles_list"].apply(_blend_role_color)
    table["tooltip"] = table.apply(
        lambda row: f"Team {int(row['Team Nr.'])} – {row['role']}", axis=1
    )

    points = table[["Team Nr.", "longitude", "latitude", "role", "color", "tooltip"]].copy()
    points["longitude"] = points["longitude"].astype(float)
    points["latitude"] = points["latitude"].astype(float)

    edges: List[Dict[str, Any]] = []
    stage_definitions = [
        ("Zur Vorspeise", "Sees 1", starters_set),
        ("Zur Hauptspeise", "Sees 3", main_set),
        ("Zum Dessert", "Sees 5", dessert_set),
    ]

    for _, row in table.iterrows():
        team_id = int(row["Team Nr."])
        for stage_name, sees_column, host_set in stage_definitions:
            raw_target = row.get(sees_column, 0)
            if pd.isna(raw_target):
                continue
            host_id = int(raw_target)
            if host_id == 0 or host_id == team_id:
                continue
            if team_id in host_set:
                # Hosts stay at home; the travelling teams are the guests.
                continue

            host_row = table.loc[table["Team Nr."] == host_id]
            if host_row.empty:
                continue
            host_row = host_row.iloc[0]

            edges.append(
                {
                    "team": team_id,
                    "target": host_id,
                    "stage": stage_name,
                    "color": list(STAGE_COLORS.get(stage_name, (120, 120, 120))),
                    "start": [float(row["longitude"]), float(row["latitude"])],
                    "end": [
                        float(host_row["longitude"]),
                        float(host_row["latitude"]),
                    ],
                    "tooltip": f"Team {team_id} → {host_id} ({stage_name})",
                }
            )

    edges_df = pd.DataFrame(edges)
    return points, edges_df


def build_circle_chart(points_df: pd.DataFrame, edges_df: pd.DataFrame) -> "go.Figure":
    """
    Arrange teams on a circle and connect them via lines for each dinner stage.

    Returns a Plotly figure object that can be embedded in Streamlit.
    """

    if not PLOTLY_AVAILABLE:
        raise RuntimeError("Plotly ist nicht verfügbar. Bitte installiere es mit 'pip install plotly'.")

    if points_df.empty:
        raise ValueError("Keine Teams vorhanden – Kreisdarstellung kann nicht erzeugt werden.")

    sorted_points = points_df.sort_values("Team Nr.").reset_index(drop=True)
    node_positions: Dict[int, Tuple[float, float]] = {}

    team_count = len(sorted_points)
    radius = 1.0

    for idx, row in sorted_points.iterrows():
        team_id = int(row["Team Nr."])
        angle = 2 * math.pi * idx / team_count
        x = radius * math.cos(angle)
        y = radius * math.sin(angle)
        node_positions[team_id] = (x, y)
        sorted_points.loc[idx, "x"] = x
        sorted_points.loc[idx, "y"] = y
        sorted_points.loc[idx, "color_hex"] = _rgb_to_hex(row["color"])

    stage_segments: Dict[str, Dict[str, List[float | Optional[str]]]] = {}
    for stage_name, stage_color in STAGE_COLORS.items():
        stage_segments[stage_name] = {
            "x": [],
            "y": [],
            "text": [],
            "color": _rgb_to_hex(stage_color),
        }

    for _, edge in edges_df.iterrows():
        stage = str(edge["stage"])
        team_id = int(edge["team"])
        host_id = int(edge["target"])

        if team_id not in node_positions or host_id not in node_positions:
            continue

        start_x, start_y = node_positions[team_id]
        end_x, end_y = node_positions[host_id]

        segments = stage_segments.setdefault(
            stage,
            {
                "x": [],
                "y": [],
                "text": [],
                "color": _rgb_to_hex(STAGE_COLORS.get(stage, (150, 150, 150))),
            },
        )
        tooltip = edge.get("tooltip", f"{team_id} → {host_id}")
        segments["x"].extend([start_x, end_x, None])
        segments["y"].extend([start_y, end_y, None])
        segments["text"].extend([tooltip, tooltip, None])

    line_traces: List["go.Scatter"] = []
    for stage, data in stage_segments.items():
        if not data["x"]:
            continue
        line_traces.append(
            go.Scatter(
                x=data["x"],
                y=data["y"],
                mode="lines",
                line=dict(color=data["color"], width=3),
                name=stage,
                hoverinfo="text",
                text=data["text"],
            )
        )

    node_trace = go.Scatter(
        x=sorted_points["x"],
        y=sorted_points["y"],
        mode="markers+text",
        marker=dict(
            color=sorted_points["color_hex"],
            size=18,
            line=dict(width=1, color="#333333"),
        ),
        text=[f"{int(row['Team Nr.'])}" for _, row in sorted_points.iterrows()],
        hoverinfo="text",
        textposition="top center",
        hovertext=[row["tooltip"] for _, row in sorted_points.iterrows()],
        name="Teams",
    )

    layout = go.Layout(
        showlegend=True,
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        margin=dict(l=20, r=20, t=40, b=20),
        height=600,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    figure = go.Figure(data=line_traces + [node_trace], layout=layout)
    figure.update_yaxes(scaleanchor="x", scaleratio=1)
    return figure


# --------------------------------------------------------------------------------------
# Main orchestration
# --------------------------------------------------------------------------------------

def plan_running_dinner(
    address_df: pd.DataFrame,
    afterparty_df: pd.DataFrame,
    api_key: str,
    config: PlannerConfig,
    *,
    cache_path: Path = GEOCODE_CACHE_PATH,
    request_pause: float = 1.0,
) -> PlannerResult:
    """Turn the raw input tables into assignments, overview text, and mail drafts."""

    _validate_columns(address_df, MANDATORY_TEAM_COLUMNS, "Adressliste")
    _validate_columns(afterparty_df, MANDATORY_AFTERPARTY_COLUMNS, "Afterparty")

    client = _get_ors_client(api_key)

    cache = _load_geocode_cache(cache_path)
    address_geocoded = _geocode_dataframe(address_df, client, cache, request_pause)
    afterparty_geocoded = _geocode_dataframe(afterparty_df, client, cache, request_pause)
    _save_geocode_cache(cache_path, cache)

    distance_df = _prepare_distance_to_afterparty(address_geocoded, afterparty_geocoded, client)

    team_count = len(distance_df)
    double_cook_quota = (3 - (team_count % 3)) % 3
    double_cook_numbers = _select_double_cooks(distance_df, double_cook_quota)
    starters, main_courses, desserts = _assign_courses(distance_df, double_cook_numbers, double_cook_quota)

    starters_table = _build_course_table(distance_df, starters)
    main_courses_table = _build_course_table(distance_df, main_courses)
    desserts_table = _build_course_table(distance_df, desserts)

    distance_star_main = _many_to_many_distances(
        client, starters_table, main_courses_table, "Starter Team Nr.", "Main Course Team Nr."
    )
    distance_main_dess = _many_to_many_distances(
        client, main_courses_table, desserts_table, "Main Course Team Nr.", "Dessert Team Nr."
    )
    distance_star_dess = _many_to_many_distances(
        client, starters_table, desserts_table, "Starter Course Team Nr.", "Dessert Team Nr."
    )

    address_table = address_geocoded.copy()
    for column in SEES_COLUMNS:
        address_table[column] = 0
    address_table.set_index("Team Nr.", inplace=True)

    # Populate encounter columns following the original order.
    _assignment(
        address_table,
        main_courses_table,
        distance_star_main,
        True,
        "Sees 1",
        "Sees 2",
        "Main Course Team Nr.",
        "Starter Team Nr.",
    )

    _assignment(
        address_table,
        desserts_table,
        distance_star_dess,
        True,
        "Sees 1",
        "Sees 2",
        "Dessert Team Nr.",
        "Starter Course Team Nr.",
    )

    copy_star_dess = _overwrite_routes(
        distance_star_dess,
        distance_star_main,
        address_table,
        "Dessert Team Nr.",
        "Starter Course Team Nr.",
        "Starter Team Nr.",
        "Main Course Team Nr.",
        "Sees 1",
    )
    copy_star_dess = copy_star_dess.sort_values(
        by=["Dessert Team Nr.", "Starter Course Team Nr."], ascending=[True, True]
    ).copy()
    copy_star_dess["Main Course Team Nr."] = copy_star_dess["Starter Course Team Nr."]
    copy_star_dess.drop(columns=["Starter Course Team Nr."], inplace=True)

    _assignment(
        address_table,
        desserts_table,
        copy_star_dess,
        False,
        "Sees 3",
        "Sees 4",
        "Dessert Team Nr.",
        "Main Course Team Nr.",
    )

    _assignment(
        address_table,
        starters_table,
        distance_star_main,
        True,
        "Sees 3",
        "Sees 4",
        "Starter Team Nr.",
        "Main Course Team Nr.",
    )

    copy_star_main = _overwrite_routes(
        distance_star_main,
        distance_main_dess,
        address_table,
        "Starter Team Nr.",
        "Main Course Team Nr.",
        "Main Course Team Nr.",
        "Dessert Team Nr.",
        "Sees 3",
    )
    copy_star_main["Dessert Team Nr."] = copy_star_main["Main Course Team Nr."]
    copy_star_main.drop(columns=["Main Course Team Nr."], inplace=True)

    _assignment(
        address_table,
        main_courses_table,
        distance_main_dess,
        True,
        "Sees 5",
        "Sees 6",
        "Main Course Team Nr.",
        "Dessert Team Nr.",
    )

    _assignment(
        address_table,
        starters_table,
        copy_star_main,
        False,
        "Sees 5",
        "Sees 6",
        "Starter Team Nr.",
        "Dessert Team Nr.",
    )

    overview_text = _build_overview_text(address_table, starters, main_courses, desserts)

    afterparty_row = afterparty_geocoded.iloc[0]
    mail_contents: Dict[int, str] = {}
    for team_id in address_table.index.astype(int):
        mail_contents[int(team_id)] = _mail_body(
            address_table,
            int(team_id),
            config,
            starters,
            main_courses,
            desserts,
            afterparty_row,
        )

    distance_frames = {
        "Distance_Starters_to_MainCourses": distance_star_main,
        "Distance_MainCourses_to_Desserts": distance_main_dess,
        "Distance_Starters_to_Desserts": distance_star_dess,
    }

    return PlannerResult(
        address_table=address_table.reset_index(),
        starters_table=starters_table,
        main_courses_table=main_courses_table,
        desserts_table=desserts_table,
        distance_frames=distance_frames,
        overview_text=overview_text,
        mail_contents=mail_contents,
        double_cook_teams=double_cook_numbers,
    )


# --------------------------------------------------------------------------------------
# Output utilities
# --------------------------------------------------------------------------------------

def write_outputs(result: PlannerResult, output_dir: Path = DEFAULT_OUTPUT_DIR) -> None:
    """Persist the overview, mail drafts, and enriched address list to disk."""

    output_dir.mkdir(parents=True, exist_ok=True)
    overview_path = output_dir / "overview.txt"
    overview_path.write_text(result.overview_text, encoding="utf-8")

    mail_dir = output_dir / "mails"
    mail_dir.mkdir(parents=True, exist_ok=True)
    for team_id, content in result.mail_contents.items():
        mail_path = mail_dir / f"Mail Team Nr.{team_id}.txt"
        mail_path.write_text(content, encoding="utf-8")

    excel_path = output_dir / "adress_list_with_assignments.xlsx"
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        result.address_table.to_excel(writer, index=False, sheet_name="Teams")
        result.starters_table.to_excel(writer, index=False, sheet_name="Starters")
        result.main_courses_table.to_excel(writer, index=False, sheet_name="MainCourses")
        result.desserts_table.to_excel(writer, index=False, sheet_name="Desserts")


def _dataframe_to_excel_bytes(frames: Dict[str, pd.DataFrame]) -> BytesIO:
    """Pack several dataframes into a single Excel workbook and return it as bytes."""

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for sheet_name, frame in frames.items():
            frame.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    buffer.seek(0)
    return buffer


def _mail_zip_bytes(mail_contents: Dict[int, str]) -> BytesIO:
    """Return a ZIP archive containing all mail drafts."""

    buffer = BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for team_id, content in mail_contents.items():
            archive.writestr(f"Mail Team Nr.{team_id}.txt", content)
    buffer.seek(0)
    return buffer


# --------------------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------------------

def run_streamlit_app() -> None:
    """Launch the Streamlit UI if the library is available."""

    if not STREAMLIT_AVAILABLE:
        raise RuntimeError(
            "Streamlit ist nicht installiert. Bitte führe 'pip install streamlit' aus."
        )

    st.set_page_config(page_title="Running Dinner Planner", layout="wide")
    st.title("Running Dinner Planner")
    st.caption("Erstellt Kurszuweisungen, Überblick und Mailtexte für euer Running Dinner.")

    with st.expander("Kurzanleitung", expanded=False):
        st.markdown(
            """
            1. Trage deinen OpenRouteService-API-Key ein.
            2. Lade `data.xlsx` (Adressliste) und `Afterparty.xlsx` hoch.
            3. Fülle Zeiten und Kontakte aus.
            4. Klicke auf **Plan erstellen** und warte bis die Distanzberechnungen abgeschlossen sind.
            """
        )

    api_key = st.text_input("OpenRouteService API-Key", type="password")
    data_file = st.file_uploader("Adressliste (`data.xlsx`)", type=["xlsx"])
    afterparty_file = st.file_uploader("Afterparty (`Afterparty.xlsx`)", type=["xlsx"])

    default_language = "Deutsch"
    language_map = {"Deutsch": "de", "English": "en"}

    with st.form("planner_config"):
        language_choice = st.selectbox("Sprache für Mails", options=list(language_map.keys()), index=0)
        time_starter = st.text_input("Startzeit Vorspeise", value="18:00")
        time_main = st.text_input("Startzeit Hauptspeise", value="20:00")
        time_dessert = st.text_input("Startzeit Nachspeise", value="22:00")
        time_afterparty = st.text_input("Startzeit Afterparty", value="23:30")
        drinks_option = st.selectbox(
            "Afterparty Getränke",
            options=["Spendenbasis", "Kostenlos bereitgestellt"],
            index=0,
        )
        organizer_name = st.text_input("Orga-Ansprechpartner*in", value="")
        organizer_phone = st.text_input("Telefonnummer Orga", value="")
        awareness_enabled = st.checkbox("Awareness-Team vorhanden?")
        awareness_phone = st.text_input("Telefonnummer Awareness-Team", value="")
        submitted = st.form_submit_button("Plan erstellen")

    if submitted:
        if not api_key:
            st.error("Bitte trage deinen OpenRouteService API-Key ein.")
            return
        if data_file is None or afterparty_file is None:
            st.error("Bitte lade sowohl `data.xlsx` als auch `Afterparty.xlsx` hoch.")
            return

        config = PlannerConfig(
            language=language_map.get(language_choice, "de"),
            time_starter=time_starter.strip() or "18:00",
            time_main=time_main.strip() or "20:00",
            time_dessert=time_dessert.strip() or "22:00",
            time_afterparty=time_afterparty.strip() or "23:30",
            drinks_option="paid" if drinks_option.startswith("Spenden") else "free",
            organizer_name=organizer_name.strip(),
            organizer_phone=organizer_phone.strip(),
            awareness_enabled=awareness_enabled,
            awareness_phone=awareness_phone.strip() or None,
        )

        with st.spinner("Berechne Routen und Kurszuweisungen ..."):
            try:
                address_df = _read_excel(data_file)
                afterparty_df = _read_excel(afterparty_file)
                result = plan_running_dinner(address_df, afterparty_df, api_key, config)
            except Exception as exc:  # pylint: disable=broad-except - we show the exception to the user
                st.error(f"Planung fehlgeschlagen: {exc}")
                st.exception(exc)
                return

        st.success("Planung erfolgreich abgeschlossen!")
        st.write(f"Doppel-Koch-Teams: {', '.join(map(str, result.double_cook_teams)) or 'keine benötigt'}")

        try:
            points_df, edges_df = prepare_map_frames(
                result.address_table,
                result.starters_table["Team Nr."].astype(int).tolist(),
                result.main_courses_table["Team Nr."].astype(int).tolist(),
                result.desserts_table["Team Nr."].astype(int).tolist(),
            )
        except Exception as exc:
            st.warning(f"Visualisierung konnte nicht vorbereitet werden: {exc}")
            points_df = pd.DataFrame()
            edges_df = pd.DataFrame()

        if PLOTLY_AVAILABLE and not points_df.empty:
            try:
                st.subheader("Dinner-Netzwerk (Kreislayout)")
                circle_fig = build_circle_chart(points_df, edges_df)
                st.plotly_chart(circle_fig, use_container_width=True)
                st.caption(
                    "Alle Teams sind gleichmäßig auf einem Kreis angeordnet. "
                    "Kreisfarbe zeigt die Gastgeberrolle (blau = Starter, grün = Main, koralle = Dessert). "
                    "Linienfarben: orange → Vorspeise, türkis → Hauptspeise, grün → Dessert."
                )
            except Exception as exc:
                st.warning(f"Kreisdarstellung konnte nicht erzeugt werden: {exc}")
        elif not PLOTLY_AVAILABLE:
            st.info("Für die Kreisdarstellung bitte zusätzlich `pip install plotly` ausführen.")

        if PYDECK_AVAILABLE and not points_df.empty:
            with st.expander("Geografische Karte (optional)"):
                center_lon = float(points_df["longitude"].mean())
                center_lat = float(points_df["latitude"].mean())
                layers = [
                    pdk.Layer(
                        "ScatterplotLayer",
                        data=points_df,
                        get_position="[longitude, latitude]",
                        get_fill_color="color",
                        get_radius=90,
                        pickable=True,
                        radius_min_pixels=4,
                    )
                ]

                if not edges_df.empty:
                    layers.append(
                        pdk.Layer(
                            "LineLayer",
                            data=edges_df,
                            get_source_position="start",
                            get_target_position="end",
                            get_color="color",
                            get_width=4,
                            pickable=True,
                        )
                    )

                deck = pdk.Deck(
                    map_style="mapbox://styles/mapbox/light-v9",
                    initial_view_state=pdk.ViewState(
                        latitude=center_lat,
                        longitude=center_lon,
                        zoom=12,
                    ),
                    layers=layers,
                    tooltip={"html": "<b>{tooltip}</b>", "style": {"color": "white"}},
                )

                st.pydeck_chart(deck, use_container_width=True)
                st.caption(
                    "Die Karte zeigt reale Standorte. "
                    "Punkte visualisieren Gastgeberrollen, Linien markieren die Wege je Kurs."
                )
        elif not PYDECK_AVAILABLE:
            st.info("Für die optionale Kartenansicht bitte zusätzlich `pip install pydeck` ausführen.")

        st.subheader("Download")
        st.download_button(
            "overview.txt herunterladen",
            data=result.overview_text,
            file_name="overview.txt",
            mime="text/plain",
        )
        st.download_button(
            "Mail-Entwürfe (ZIP)",
            data=_mail_zip_bytes(result.mail_contents),
            file_name="mail_entwuerfe.zip",
            mime="application/zip",
        )

        excel_frames = {
            "Teams": result.address_table,
            "Starters": result.starters_table,
            "MainCourses": result.main_courses_table,
            "Desserts": result.desserts_table,
        }
        st.download_button(
            "Excel mit Zuweisungen",
            data=_dataframe_to_excel_bytes(excel_frames),
            file_name="running_dinner_assignments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Tabellen")
        st.markdown("**Adressliste mit Begegnungen**")
        st.dataframe(result.address_table, use_container_width=True)
        st.markdown("**Starter Hosts**")
        st.dataframe(result.starters_table, use_container_width=True)
        st.markdown("**Main Course Hosts**")
        st.dataframe(result.main_courses_table, use_container_width=True)
        st.markdown("**Dessert Hosts**")
        st.dataframe(result.desserts_table, use_container_width=True)


# --------------------------------------------------------------------------------------
# CLI handling
# --------------------------------------------------------------------------------------

def _load_config(path: Path) -> PlannerConfig:
    """Read organiser settings from JSON."""

    if not path.exists():
        raise FileNotFoundError(f"Die Konfigurationsdatei {path} wurde nicht gefunden.")

    with path.open("r", encoding="utf-8") as handle:
        raw = json.load(handle)

    return PlannerConfig(
        language=raw.get("language", "de"),
        time_starter=raw.get("time_starter", "18:00"),
        time_main=raw.get("time_main", "20:00"),
        time_dessert=raw.get("time_dessert", "22:00"),
        time_afterparty=raw.get("time_afterparty", "23:30"),
        drinks_option=raw.get("drinks_option", "paid"),
        organizer_name=raw.get("organizer_name", ""),
        organizer_phone=raw.get("organizer_phone", ""),
        awareness_enabled=raw.get("awareness_enabled", False),
        awareness_phone=raw.get("awareness_phone"),
    )


def run_cli() -> None:
    """Command-line entry point."""

    parser = argparse.ArgumentParser(description="Running Dinner Planner (CLI)")
    parser.add_argument("--addresses", default="data.xlsx", help="Pfad zur Adressliste (Excel)")
    parser.add_argument("--afterparty", default="Afterparty.xlsx", help="Pfad zur Afterparty-Datei (Excel)")
    parser.add_argument(
        "--config",
        default="config.json",
        help="Pfad zur JSON-Konfiguration (Zeiten, Kontakte, Sprache)",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Ausgabeordner für overview, Mails und Excel",
    )
    parser.add_argument(
        "--pause",
        type=float,
        default=1.0,
        help="Wartezeit (in Sekunden) zwischen Geocoding-Anfragen",
    )
    parser.add_argument(
        "--api-key",
        default=os.environ.get("ORS_API_KEY", ""),
        help="OpenRouteService API-Key (überschreibt ORS_API_KEY)",
    )

    args = parser.parse_args()

    config = _load_config(Path(args.config))
    api_key = args.api_key or os.environ.get("ORS_API_KEY", "")

    address_df = _read_excel(args.addresses)
    afterparty_df = _read_excel(args.afterparty)

    result = plan_running_dinner(
        address_df,
        afterparty_df,
        api_key,
        config,
        cache_path=GEOCODE_CACHE_PATH,
        request_pause=args.pause,
    )
    write_outputs(result, Path(args.output))
    print(f"Planung abgeschlossen. Ergebnisse wurden in '{args.output}' abgelegt.")


# --------------------------------------------------------------------------------------
# Dispatch depending on environment
# --------------------------------------------------------------------------------------

def _running_inside_streamlit() -> bool:
    """Detect whether the module is executed via `streamlit run`."""

    if not STREAMLIT_AVAILABLE:
        return False
    try:
        return get_script_run_ctx() is not None
    except Exception:  # pragma: no cover - Streamlit internals might change
        return False


if _running_inside_streamlit():
    run_streamlit_app()
elif __name__ == "__main__":
    run_cli()

