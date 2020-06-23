from selenium.webdriver.chrome.options import Options


class WebDriverMonkeyPatches:

    @property
    def experimental_options(self):
        """
        Returns a dictionary of experimental options for chrome.
        """
        if self._experimental_options:
            return self._experimental_options
        else:
            return {"prefs": {"download.prompt_for_download": True}}

    Options.experimental_options = experimental_options
