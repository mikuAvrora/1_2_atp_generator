import json


class Project:
    @property
    def title(self, *args, **kwargs):
        with open('config/config.json', 'r', encoding="utf-8") as file:
            return json.load(file)["title"]
        
    @property
    def show_errors_window(self, *args, **kwargs):
        with open('config/config.json', 'r', encoding="utf-8") as file:
            return json.load(file)["show_errors_window"]
        
    @property
    def show_warning(self, *args, **kwargs):
        with open('config/config.json', 'r', encoding="utf-8") as file:
            return json.load(file)["show_warning"]
