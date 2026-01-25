import json
from pathlib import Path

from .constants import SCHEMA_VERSION
from .model import ProjectModel


def save_project(path: str | Path, model: ProjectModel, view_state: dict) -> None:
    payload = model.to_dict()
    payload["schema_version"] = SCHEMA_VERSION
    payload["view"] = view_state
    path = Path(path)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_project(path: str | Path) -> tuple[ProjectModel, dict]:
    path = Path(path)
    data = json.loads(path.read_text(encoding="utf-8"))
    model = ProjectModel.from_dict(data)
    view_state = data.get("view", {})
    return model, view_state
