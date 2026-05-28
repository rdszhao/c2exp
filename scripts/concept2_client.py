"""
Concept2 Logbook API Client

Handles OAuth2 authentication and fetching workout data from Concept2 Logbook.
"""

import json
import re
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlencode, urlparse, parse_qs
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import requests


BASE_URL = "https://log.concept2.com"
AUTH_URL = f"{BASE_URL}/oauth/authorize"
TOKEN_URL = f"{BASE_URL}/oauth/access_token"
DATA_DIR = Path(__file__).parent.parent / "data"
TOKEN_FILE = DATA_DIR / ".c2_token.json"


@dataclass
class Stroke:
    """Single stroke data point."""
    index: int
    power: float


@dataclass
class Workout:
    """Represents a single workout with stroke data."""
    date: str
    person: str
    side: str  # L or R
    strokes: list[Stroke] = field(default_factory=list)
    avg_power: Optional[float] = None
    peak_power: Optional[float] = None
    peak_index: Optional[int] = None

    @staticmethod
    def parse_notes(notes: str) -> tuple[str, str]:
        """Extract person name and side (L/R) from notes.

        Expected format: 'Name L' or 'Name R'
        """
        if not notes:
            return "Unknown", ""

        notes = notes.strip()
        # Match pattern like "John L" or "Jane R" at start of notes
        match = re.match(r'^(.+?)\s+([LR])\b', notes, re.IGNORECASE)
        if match:
            return match.group(1).strip(), match.group(2).upper()

        # Try just ending with L or R
        if notes.endswith(' L') or notes.endswith(' l'):
            return notes[:-2].strip(), 'L'
        if notes.endswith(' R') or notes.endswith(' r'):
            return notes[:-2].strip(), 'R'

        return notes, ""


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    """HTTP handler to capture OAuth callback."""

    def do_GET(self):
        query = parse_qs(urlparse(self.path).query)
        if "code" in query:
            self.server.auth_code = query["code"][0]
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()
            self.wfile.write(b"<html><body><h1>Authorization successful!</h1>"
                           b"<p>You can close this window.</p></body></html>")
        else:
            self.server.auth_code = None
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"Authorization failed")

    def log_message(self, format, *args):
        pass


class Concept2Client:
    """Client for the Concept2 Logbook API."""

    def __init__(self, client_id: str, client_secret: str, redirect_uri: str):
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.access_token: Optional[str] = None
        self.refresh_token: Optional[str] = None
        self.user_id: Optional[int] = None
        self._load_token()

    def _load_token(self):
        if TOKEN_FILE.exists():
            try:
                data = json.loads(TOKEN_FILE.read_text())
                self.access_token = data.get("access_token")
                self.refresh_token = data.get("refresh_token")
                self.user_id = data.get("user_id")
            except (json.JSONDecodeError, KeyError):
                pass

    def _save_token(self):
        TOKEN_FILE.write_text(json.dumps({
            "access_token": self.access_token,
            "refresh_token": self.refresh_token,
            "user_id": self.user_id,
        }))

    def authenticate(self) -> bool:
        if self.access_token and self._test_token():
            return True

        params = {
            "client_id": self.client_id,
            "redirect_uri": self.redirect_uri,
            "response_type": "code",
        }
        auth_url = f"{AUTH_URL}?{urlencode(params)}"

        parsed = urlparse(self.redirect_uri)
        port = parsed.port or 8080

        server = HTTPServer(("localhost", port), OAuthCallbackHandler)
        server.auth_code = None

        print(f"Opening browser for authorization...")
        webbrowser.open(auth_url)

        server.handle_request()

        if not server.auth_code:
            print("Authorization failed - no code received")
            return False

        response = requests.post(TOKEN_URL, data={
            "grant_type": "authorization_code",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "redirect_uri": self.redirect_uri,
            "code": server.auth_code,
        })

        if response.status_code != 200:
            print(f"Token exchange failed: {response.text}")
            return False

        data = response.json()
        self.access_token = data["access_token"]
        self.refresh_token = data.get("refresh_token")

        if not self._fetch_user_id():
            return False

        self._save_token()
        print("Authentication successful!")
        return True

    def _test_token(self) -> bool:
        if not self.access_token:
            return False
        try:
            resp = self._api_get("/api/users/me")
            return resp is not None
        except:
            return False

    def _fetch_user_id(self) -> bool:
        data = self._api_get("/api/users/me")
        if data and "data" in data:
            self.user_id = data["data"]["id"]
            return True
        return False

    def _api_get(self, endpoint: str, params: dict = None) -> Optional[dict]:
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(f"{BASE_URL}{endpoint}", headers=headers, params=params)
        if response.status_code == 200:
            return response.json()
        return None

    def get_workouts(self) -> list[Workout]:
        """Fetch all workouts with full stroke data."""
        if not self.user_id:
            if not self._fetch_user_id():
                raise RuntimeError("Could not get user ID")

        workouts = []
        page = 1

        while True:
            print(f"Fetching workouts page {page}...")
            data = self._api_get(
                f"/api/users/{self.user_id}/results",
                params={"per_page": 250, "page": page}
            )

            if not data or "data" not in data:
                break

            for result in data["data"]:
                workout = self._parse_workout(result)
                if workout:
                    workouts.append(workout)

            meta = data.get("meta", {}).get("pagination", {})
            if page >= meta.get("total_pages", 1):
                break
            page += 1

        print(f"Fetched {len(workouts)} workouts")
        return workouts

    def _parse_workout(self, data: dict) -> Optional[Workout]:
        try:
            notes = data.get("comments", "") or ""
            person, side = Workout.parse_notes(notes)

            workout = Workout(
                date=data.get("date", ""),
                person=person,
                side=side,
            )

            # Fetch stroke data
            self._load_strokes(data["id"], workout)

            return workout
        except (KeyError, TypeError) as e:
            print(f"Error parsing workout: {e}")
            return None

    def _load_strokes(self, result_id: int, workout: Workout):
        """Load stroke data and compute power metrics."""
        if not self.user_id:
            return

        data = self._api_get(f"/api/users/{self.user_id}/results/{result_id}/strokes")
        if not data or "data" not in data:
            return

        strokes_data = data["data"]
        if not strokes_data:
            return

        powers = []
        for i, stroke in enumerate(strokes_data):
            pace_tenths = stroke.get("p", 0)
            if pace_tenths and pace_tenths > 0:
                pace_seconds = pace_tenths / 10.0
                power = 2.80 / ((pace_seconds / 500.0) ** 3)
                power = round(power, 1)
                workout.strokes.append(Stroke(index=i + 1, power=power))
                powers.append((i, power))

        if powers:
            workout.avg_power = round(sum(p for _, p in powers) / len(powers), 1)
            peak_idx, peak_val = max(powers, key=lambda x: x[1])
            workout.peak_power = peak_val
            workout.peak_index = peak_idx + 1  # 1-indexed
