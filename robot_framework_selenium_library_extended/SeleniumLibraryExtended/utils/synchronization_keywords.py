from SeleniumLibrary import SeleniumLibrary
import robot
import time
from robot.api import logger

class _SynchronizationKeywords(SeleniumLibrary):
    def _get_visible_item(self, locator):
        elements = self._element_finder.find(locator,first_only=False)
        for element in elements:
            if self._is_element_present(element) and self._is_visible(element):
                return element

    def _wait_and_get_visible_item(self, locator, timeout=None):
        error = "Element '%s' was not visible in %s" % (locator, self._format_timeout(timeout))
        #self._wait_until(timeout, error, self._get_visible_item, locator)
        self._wait_until(lambda: self._get_visible_item(locator) is not None, error, timeout)
        return self._get_visible_item(locator)

    def _wait_for_jquery(self):
        self.wait_for_condition("return Boolean(window.jQuery);")

    def _is_element_present(self, locator, tag=None):
        return (self._element_find(locator, True, False, tag=tag) is not None)

    def _format_timeout(self, timeout):
        timeout = robot.utils.timestr_to_secs(timeout) if timeout is not None else self.timeout
        return robot.utils.secs_to_timestr(timeout)

    def _is_visible(self, locator):
        element = self._element_find(locator, required=False)
        return element.is_displayed() if element else None

    def _wait_until_no_error(self, timeout, wait_func, *args):
        timeout = robot.utils.timestr_to_secs(timeout) if timeout is not None else self.timeout
        maxtime = time.time() + timeout
        while True:
            timeout_error = wait_func(*args)
            if not timeout_error: return
            if time.time() > maxtime:
                raise AssertionError(timeout_error)
            time.sleep(0.2)

    def _element_find(self, locator, first_only=True, required=True, tag=None):
        return self._element_finder.find(locator, tag, first_only, required)

    def _current_browser(self):
        return self.driver

    def _get_text_of_all_elements(self, locator, timeout=10):
        self.wait_until_page_contains_element(locator, timeout)
        self.wait_until_element_is_enabled(locator, timeout)
        elementlist = self.get_webelements(locator)
        list_i = [self.get_text(e) for e in elementlist]
        return list_i

    def _get_value(self, locator, tag):
        return self.find_element(locator, tag).get_attribute('value')

    @property
    def _cache(self):
        return self._drivers

    def _info(self, message, html=False, also_console=False):
        logger.info(message,html,also_console)

    def _log(self, message):
        logger.info(message)

    def _debug(self, message):
        logger.debug(message)

    def _warn(self, message):
        logger.warn(message)

    def _html(self, message):
        logger.info(message, True, False)






