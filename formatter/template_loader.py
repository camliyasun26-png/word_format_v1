# formatter/template_loader.py
import os
import yaml

def load_template(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)

def list_templates(directory: str) -> list[str]:
    if not os.path.isdir(directory):
        return []
    return sorted(f for f in os.listdir(directory) if f.endswith(".yaml"))
