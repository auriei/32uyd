import json
import os

class ConfigManager:
    def __init__(self, config_dir='data/config'):
        self.config_dir = config_dir
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir, exist_ok=True) # Added exist_ok=True

    def get_config_filepath(self, config_name):
        if not config_name.endswith('.json'):
            config_name += '.json'
        return os.path.join(self.config_dir, config_name)

    def load_config(self, config_name, default_settings=None):
        config_filepath = self.get_config_filepath(config_name)
        
        if not os.path.exists(config_filepath):
            current_settings = default_settings if default_settings is not None else {}
            self.save_config(config_name, current_settings) # Save defaults if file doesn't exist
            return current_settings
        
        try:
            with open(config_filepath, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            # If default_settings are provided, ensure all keys are present
            if default_settings is not None:
                updated = False
                for key, value in default_settings.items():
                    if key not in settings:
                        settings[key] = value
                        updated = True
                if updated:
                    self.save_config(config_name, settings) # Save updated settings
            return settings
        except (IOError, json.JSONDecodeError) as e:
            print(f"Error loading {config_name} from {config_filepath}: {e}. Retrying with defaults.")
            current_settings = default_settings if default_settings is not None else {}
            self.save_config(config_name, current_settings) # Attempt to save defaults
            return current_settings

    def save_config(self, config_name, settings):
        config_filepath = self.get_config_filepath(config_name)
        try:
            # Ensure directory exists one more time, in case it was deleted
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir, exist_ok=True)
            with open(config_filepath, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except IOError as e:
            print(f"Error saving {config_name} to {config_filepath}: {e}")

if __name__ == '__main__':
    # Example Usage
    # This example assumes the script is run from the project root directory
    # For actual use, paths might need adjustment based on where ConfigManager is instantiated
    
    # Create a ConfigManager instance. If 'data/config' doesn't exist, it will be created.
    config_manager = ConfigManager(config_dir='data/config')

    # Define default settings for app_settings
    default_app_settings = {
        "window_size": [1024, 768],
        "default_pdf_input_dir": "pdf_documents", # Corrected path relative to project root
        "default_excel_output_dir": "data/exports", # Corrected path relative to project root
        "theme": "light"
    }
    
    print(f"Attempting to load app_settings.json from {os.path.join(config_manager.config_dir, 'app_settings.json')}")
    
    # Load app_settings.json, creating it with defaults if it doesn't exist
    app_settings = config_manager.load_config('app_settings.json', default_settings=default_app_settings)
    print(f"Loaded app_settings: {app_settings}")

    # Modify a setting
    app_settings['theme'] = 'dark'
    app_settings['new_setting_test'] = 'test_value'
    config_manager.save_config('app_settings.json', app_settings)
    print(f"Saved app_settings: {app_settings}")

    reloaded_app_settings = config_manager.load_config('app_settings.json', default_settings=default_app_settings)
    print(f"Reloaded app_settings (should include new_setting_test and defaults): {reloaded_app_settings}")

    # Example for another config file
    default_system_settings = {
        "backup_enabled": True,
        "backup_interval_hours": 24,
        "log_level": "INFO"
    }
    system_settings = config_manager.load_config('system_settings', default_settings=default_system_settings) # .json added automatically
    print(f"Loaded system_settings: {system_settings}")
    
    # Test loading a config file that might not exist, with no defaults provided initially by load_config
    # but where defaults are applied because the file is created for the first time.
    test_settings_no_defaults_on_load = config_manager.load_config('test_settings_new.json', default_settings={'key1': 'value1'})
    print(f"Loaded test_settings_new.json (created with defaults): {test_settings_no_defaults_on_load}")

    # Test loading an existing config file but providing new defaults (should merge)
    app_settings_new_defaults = config_manager.load_config('app_settings.json', default_settings={
        "another_new_default_key": "default_value",
        "theme": "blue" # This default for 'theme' will be ignored as 'theme' already exists
    })
    print(f"Reloaded app_settings with new defaults provided (should merge): {app_settings_new_defaults}")
