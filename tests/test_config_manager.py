import os
import json
import pytest
from src.utils.config_manager import ConfigManager

@pytest.fixture
def manager_with_fake_path(mocker):
    """Patcha il path del config per usare una posizione sicura nei test."""
    mocker.patch("src.utils.config_manager.os.path.join", return_value="config_programma.json")
    return ConfigManager()

def test_config_manager_load_defaults(fs, manager_with_fake_path):
    manager_with_fake_path.load()
    assert manager_with_fake_path.get("rinomina_password") == "coemi"

def test_config_manager_save_and_load(fs, manager_with_fake_path):
    new_settings = {"rinomina_password": "custom_password", "email_subject": "Test"}
    
    manager_with_fake_path.save(new_settings)
    
    # Verifica ricaricando
    new_manager = ConfigManager()
    new_manager.load()
    assert new_manager.get("rinomina_password") == "custom_password"

def test_config_manager_invalid_json(fs, manager_with_fake_path):
    with open("config_programma.json", "w") as f:
        f.write("{ invalid json")
    
    manager_with_fake_path.load()
    assert manager_with_fake_path.get("rinomina_password") == "coemi"
