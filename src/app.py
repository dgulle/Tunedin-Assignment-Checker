"""Flask web application for the Intune Assignment Checker."""

import os
from flask import Flask, jsonify, render_template
from dotenv import load_dotenv
from graph_client import GraphClient

load_dotenv()

app = Flask(__name__)

# ── Configuration ────────────────────────────────────────────────────────────

TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")


def _get_client():
    """Create a Graph client, raising a clear error if credentials are missing."""
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise RuntimeError(
            "Missing Azure AD credentials. "
            "Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET "
            "environment variables."
        )
    return GraphClient(TENANT_ID, CLIENT_ID, CLIENT_SECRET)


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    """Serve the main single-page application."""
    return render_template("index.html")


@app.route("/api/groups")
def api_groups():
    """Return all Entra ID groups."""
    try:
        client = _get_client()
        groups = client.get_groups()
        return jsonify(groups)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/groups/<group_id>/assignments")
def api_group_assignments(group_id):
    """Return all Intune assignments for a specific group."""
    try:
        client = _get_client()
        data = {
            "configurations": client.get_device_configurations(group_id),
            "settingsCatalog": client.get_settings_catalog(group_id),
            "applications": client.get_applications(group_id),
            "scripts": client.get_scripts(group_id),
            "remediations": client.get_remediations(group_id),
        }
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)
