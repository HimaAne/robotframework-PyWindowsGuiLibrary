from robot.libraries.BuiltIn import BuiltIn
from robot.api import logger
import pywinauto
import pywinauto.mouse
import time
import sys
import re
import pyautogui
import os
from datetime import datetime
from datetime import timedelta
import subprocess


class PyWindowsGuiLibrary:

    """
        __Author__ = *Himavanthudu Bodapati*

        The PyWindowsGuiLibrary is a library that enables users to create tests for the windows based GUI applications.
        %TOC%
        = Open an application / Connect an existing application =

        == Opening an application ==

        Firstly *make sure application is launched*.

        To launch application, use keyword called `Launch Application`. After the Application is started,
        you can connect to the application using the keyword `Focus Application Window`.

        == Connecting an existing application ==

        This library provides support even for opened applications. *To connect an existing application
        use `Focus Application Window` without using `Launch Application` *.

        *Note:* To open a new instance of an application from given path use `Launch Application` or
        `Focus Application Window` to connect to a instance that is already opened.

        = Property Section =

        Initially to know about the application tittle property use AutoIT inspector.

        To get the Controls and properties use `Print Current Window Page Object Properties` and
        `Print Specific Object Properties On Current Window`

        *Example*

        *Notepad* application controls:

        Control Identifiers:

            Dialog - 'Untitled - Notepad*'    (L-9, T-9, R1929, B1039)
            ['Dialog', 'Untitled - NotepadDialog*', 'Untitled - Notepad*']
            *child_window(title="Untitled - Notepad*", control_type="Window")*
               |
               | Edit - 'Text Editor'    (L0, T54, R1920, B1001)
               | ['Edit', 'Edit0', 'Edit1']
               | child_window(title="Text Editor", auto_id="15", control_type="Edit")
               |    |
               |    | ScrollBar - 'Vertical'    (L1899, T54, R1920, B980)
               |    | ['Vertical', 'ScrollBar', 'VerticalScrollBar', 'ScrollBar0', 'ScrollBar1']
               |    | child_window(title="Vertical", auto_id="NonClientVerticalScrollBar", control_type="ScrollBar")
               |    |    |
               |    |    | Button - 'Line up'    (L1899, T54, R1920, B75)
               |    |    | ['Line up', 'Line upButton', 'Button', 'Button0', 'Button1']
               |    |    | child_window(title="Line up", auto_id="UpButton", control_type="Button")
               |    |    |
               |    |    | Button - 'Line down'    (L1899, T959, R1920, B980)
               |    |    | ['Line downButton', 'Line down', 'Button2']
               |    |    | child_window(title="Line down", auto_id="DownButton", control_type="Button")
               |    |
               |    | ScrollBar - 'Horizontal'    (L0, T980, R1899, B1001)
               |    | ['Horizontal', 'ScrollBar2', 'HorizontalScrollBar']
               |    | child_window(title="Horizontal", auto_id="NonClientHorizontalScrollBar", control_type="ScrollBar")
               |    |    |
               |    |    | Button - 'Column left'    (L0, T980, R21, B1001)
               |    |    | ['Button3', 'Column leftButton', 'Column left']
               |    |    | child_window(title="Column left", auto_id="UpButton", control_type="Button")
               |    |    |
               |    |    | Button - 'Column right'    (L1878, T980, R1899, B1001)
               |    |    | ['Button4', 'Column rightButton', 'Column right']
               |    |    | child_window(title="Column right", auto_id="DownButton", control_type="Button")
               |    |
               |    | Thumb - ''    (L1899, T980, R1920, B1001)
               |    | ['Thumb']
               |
               | ...

        == Window property ==
        This property is used to connect window. Probably window property will be same as window name visible in
        application. For example refer above `Property Section`.

        Below Syntax/Pattern must be follow to pass the value for window property.
        As of now Library supports Title and class name as window property.

        Pattern to pass the Window property.

         *Example if Title is available:---> (title:Untitled - Notepad)*

         *Example if class_name is available:---> (class_name:PQR)*

        == Element property ==
        *Locators*

        To find exact element property/control, analyse the hierarchy and find from printed controls identifiers.

        There is no specific syntax/pattern to pass the locator values, directly pass the matching properties to the
        `keywords`.

        *Example: Notepad edit property is:['Edit', 'Edit0', 'Edit1'], any one property is enough to use as locator.*

        == Child window ==
        Child windows can be used as property to perform certain actions like click.

        = Screenshots (on failure) =

        The PyWindowsGuiLibrary offers an option for automatic screenshots on error.
        By default this option is enabled, use keyword `Disable Screenshots On Failure` to skip the screenshot
        functionality. Alternatively, this option can be set with the ``screenshots_on_error`` argument when `importing`
        the library.

        = Multiple applications support =

        This library enables and supports multiple window based application automation as part of E2E scenarios.

        = Multiple Instances Support =
        This library handles and supports Multiple instances of same application. Index argument helps us to handle
        multiple instance where having same `window property`.

        = Sample Scripts =


        To switch between applications use `Focus Application Window` to change the controls towards the
        respective application.

        *Note:*

        There is no need to launch application everytime when you shift the application in E2E scenarios.

        Find below example testcase which covers `Multiple applications support`

        == *** Settings *** ==

        Library----PyWindowsGuiLibrary----screenshots_on_error=False

        == *** Variables *** ==

        ${App1}----C:\\Windows\\system32\\notepad.exe

        ${App2}----C:\\Program Files\\your_application.exe

        ${ABCB}----PQRS

        == *** Test Cases *** ==

        *Automation Test Case*

            ----Launch Application----${App1}

            ----Wait Until Window Present----title:Untitled - Notepad----30

            ----Focus Application Window----title:Untitled - Notepad

            ----Maximize Application Window

            ----Capture App Screenshot

            ----${t}=----Print Current Window Page Object Properties

            ----log----${t}

            ----Input Text----Edit----ABCD

            ----Minimize Application Window

            ----launch application----${App2}

            ----Wait Until Window Present----title:your_application

            ----Focus Application Window----title:your_application

            ----Maximize Application Window

            ----Click On Element----LoginButton

            ----Focus Application Window----title:*Untitled - Notepad

            ----Press Keys----{ENTER}${ABCB}

            ----Capture App Screenshot----ABCDPQRS----JPG

            ----Clear Edit Field----Edit

            ----Capture App Screenshot----Cleared

            ----Focus Application Window----title:your application

            ----Set Text---- Edit1----PQRS

            ----Clear Edit Field----Edit1

    """
    __version__ = '1.0'
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def __init__(self, screenshots_on_error=True):
        """
        Sets default variables for the library.

        | *Arguments*                | *Documentation*                 |
        | screenshots_on_error=True  | Enables screenshots on Failures |
        | screenshots_on_error=False | Disables screenshots on Failures |

        """
        self.take_screenshots = screenshots_on_error

    def disable_screenshots_on_failure(self):
        """
        Disables automatic screenshots on error.

        *Example*:
        | *Keyword*                      |                                                 |
        | Disable Screenshots On Failure | # This disables screenshots on failures/errors. |

        """
        self.take_screenshots = False

    def enable_screenshots_on_failure(self):
        """
        Enables automatic screenshots on error.

        *Example*:
        | *Keyword*                     |                                                |
        | Enable Screenshots On Failure | # This enables screenshots on failures/errors. |

        """
        self.take_screenshots = True

    def launch_application(self, path, backend="uia"):
        """
        Launches The Application from Given path.

        By default applications will open with "uia" backend process. This keyword supports both uia and win32 backend
        process.

        If Required to open the Application in win32 than need to pass, backend process as win32.

        *Example*:
        | *Keyword*          | *Attributes*               |       |                                               |
        | Launch Application | C:\\Program Files\\app.exe |       | # By default application launches with uia.   |
        | Launch Application | C:\\Program Files\\app.exe | win32 | # By default application launches with win32. |

        """
        logger.info("Launching the application from the path " + path)
        try:
            global APPInstance
            APPInstance = pywinauto.application.Application(backend=backend).start(path, timeout=60)
            return APPInstance
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to start/Launch Given Application From the path " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def _window_handle_(self, window_property, index):
        global w_handle
        if 'title:' in window_property:
            win_prop = window_property.replace('title:', '')
            w_handle = pywinauto.findwindows.find_windows(title=win_prop, ctrl_index=None, top_level_only=True,
                                                          visible_only=True)[index]
        elif 'class_name:' in window_property:
            win_prop = window_property.replace('class_name:', '')
            w_handle = pywinauto.findwindows.find_windows(class_name=win_prop, ctrl_index=None, top_level_only=True,
                                                          visible_only=True)[index]

    def focus_application_window(self, window_property, backend="uia", index=0):
        """
        This keyword Focus the Application to the Front End.

        And return the Connected Window Handle Instance :i.e dlg.

        See the `Window property` section for details about the window.

        This sets as global and the same used across further implementation.

        Use this keyword anytime to shift the control of execution between backend process win32 and uia.

        *Example*:
        | *Keyword*                | *Attributes*             | *backend process* |   | |
        | Focus Application Window | title:Untitled - Notepad |       |   | # Would focus application with backend process uia and index 0.   |
        | Focus Application Window | class_name:XYZ           | uia   | 1 | # Would focus application with backend process uia and index 1.   |
        | Focus Application Window | title:ABC                | win32 | 1 | # Would focus application with backend process win32 and index 1. |

        """
        logger.info("'Focusing the application window matches to " + window_property+"'.")
        try:
            global dlg
            global APPInstance
            APPInstance = pywinauto.application.Application(backend=backend)
            self._window_handle_(window_property, index)
            APPInstance.connect(handle=w_handle)
            dlg = APPInstance.window(handle=w_handle)
            dlg.set_focus()
            return dlg
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to Focus Application given by " + str(window_property) + " property,with " + str(index) \
                   + " index and " + str(backend) + " backend process " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def maximize_application_window(self):
        """
        Maximize the window.

        *Example*:
        | *Keyword*                   |                                     |
        | Maximize Application Window | #would maximize application window. |

        """
        logger.info("Maximizing the window")
        try:
            dlg.maximize()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to maximize the Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def minimize_application_window(self):
        """
        Minimize the window.

        *Example*:
        | *Keyword*                   |                                     |
        | Minimize Application Window | #would minimize application window. |

        """
        logger.info("minimizing the window.")
        try:
            dlg.minimize()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to minimize the Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def print_current_window_page_object_properties(self):
        """
        Returns control identifiers for current ``window``.

        *Example*:
        | *Variable*                   | *Keyword*                                  |
        | ${Complete window controls}= | #would return complete controls of window. |

        """
        logger.info("'Returns Hole window Controls From Current Active Window" + "'.")
        try:
            dlg.print_control_identifiers()
            printed_controls = sys.stdout.getvalue()
            return printed_controls
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to print and return page properties " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def print_specific_object_properties_on_current_window(self, element_property):
        """
        Returns control identifiers for given ``element property``

        See the `Element property` section for details about the locator.

        *Example*:
        | *Variable*                    | *Keyword*                                   |    *Attributes*  |  |
        | ${Specific Element controls}= | Print current window page object properties | element_property | #would return specific element controls from window. |

        """
        logger.info("Returns Specific element Controls From Given Object Property" + "'" + "``" + element_property
                    + "``" + "'.")
        try:
            dlg[element_property].print_ctrl_ids()
            printed_controls = sys.stdout.getvalue()
            return printed_controls
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to print and return Object " + element_property + " properties" + ":::: " + "because " \
                   + str(h1)
            raise Exception(mess)

    def get_text(self, element_property):
        """
        Returns the text value of the element identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Variable* | *Keyword* | *Attributes*     |                                                      |
        | ${Text}=   | Get Text  | element property | #would return text of specified element from window. |

        """
        logger.info("Retrieving element text From Given Object Property" + "'" + element_property + "'.")
        try:
            return dlg[element_property].window_text()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to Get text in Application from given by " + element_property + " :::: " + "because " \
                   + str(h1)
            raise Exception(mess)

    def set_text(self, element_property, text):
        """
        Sets the given text into the text field identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword* | *Attributes*     |      |                                                 |
        | Set Text  | element property | PQRS | #would set text to specified element in window. |

        """
        logger.info("Typing text " + text + " into text field identified by " + "'" + element_property + "'.")
        try:
            dlg[element_property].set_text(text)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to Type the given " + text + " into the text field identified by " + element_property + "." \
                   + " :::: " + "because " + str(h1)
            raise Exception(mess)

    def click_on_element(self, element_property):
        """
        Click the element identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*        | *Attributes*     |                                                      |
        | Click On Element | element property | # Would click element with by Double_Click() method. |

        """
        logger.info("Clicking element Property " + "'" + element_property + "'.")
        try:
            dlg[element_property].click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click element " + element_property + " in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def double_click_on_element(self, element_property):
        """
        Click the element identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*               | *Attributes*     |                                                      |
        | Double Click On Element | element property | # Would click element with by Double_Click() method. |

        """
        logger.info("Double Clicking element Property " + "'" + element_property + "'.")
        try:
            dlg[element_property].double_click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click object " + element_property + " in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def action_click(self, element_property):
        """
        Click the element identified by ``element_property``.

        Use Element property or child window to perform the clicking operation.

        See the `Element property` and `Child window` section for details about the locator.

        *Example*:
        | *Keyword*        | *Attributes*     |                                                     |
        | Action Click | element property | # Would click element with by Click_input() method. |
        | Action Click | child window     | # Would click element with by Click_input() method. |

        """
        logger.info("Clicking element Property " + "'" + element_property + "'.")
        try:
            dlg[element_property].click_input()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable to click object " + element_property + " in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mouseover_click(self, x_coordinates='0', y_coordinates='0', button='left'):
        """
        Click at the specified X and Y coordinates.

        By Default this performs left mouse click

        *Example*:
        | *Keyword*       | *Attributes*  |               |       |                                                    |
        | Mouseover Click |      |      |     | #By default perform left click on (0,0)    coordinates. |
        | Mouseover Click | x_coordinates | y_coordinates | right | #Would perform right click on given coordinates.   |
        | Mouseover Click | x_coordinates | y_coordinates | left | #Would perform right click on given coordinates.   |

        """
        logger.info(button + " Mouse Button clicking at X and Y coordinates i.e " + "(" + str(x_coordinates) + "," +
                    str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.click(button=button, coords=(x_coordinates, y_coordinates))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable to click object by " + str(x_coordinates), str(y_coordinates) + " in Application " + ":::: " +\
                   "because " + str(h1)
            raise Exception(mess)

    def mouseover_double_click(self, x_coordinates='0', y_coordinates='0', button='left'):
        """
        double click at the specified X and Y coordinates.

        By Default this performs left mouse double click

        *Example*:
        | *Keyword*              | *Attributes*  |               |                                                |
        | Mouseover Double Click |     |      |    | #By default perform double left click on (0,0)) coordinates. |
        | Mouseover Double Click | x_coordinates | y_coordinates | right | #Would perform right double click.   |
        | Mouseover Double Click | x_coordinates | y_coordinates | left | #Would perform left double click.   |

        """
        logger.info(button + " Mouse Button double clicking at X and Y coordinates i.e " + "(" + str(x_coordinates) +
                    "," + str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.double_click(button=button, coords=(x_coordinates, y_coordinates))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable to double click object by " + str(x_coordinates), str(y_coordinates) + " in Application " + \
                   ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mouseover_move(self, x_coordinates='0', y_coordinates='0'):
        """
        Move the mouse at specified X and Y coordinates.

        *Example*:
        | *Keyword*      | *Attributes*  |               |                                                 |
        | Mouseover move |               |               | #By default perform move on (0,0)) coordinates. |
        | Mouseover move | x_coordinates | y_coordinates | #Would perform mouse move to given coordinates. |
        """
        logger.info(" Moving mouse to given X and Y coordinates i.e " + "(" + str(x_coordinates) +
                    "," + str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.move(coords=(x_coordinates, y_coordinates))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable move mouse to coordinates " + str(x_coordinates), str(y_coordinates) + " in Application " + \
                   ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mouseover_press(self, x_coordinates='0', y_coordinates='0', button='left'):
        """
        Press the mouse button at specified X and Y coordinates.

        *Example*:
        | *Keyword*       | *Attributes*  |               |                                                 |
        | Mouseover Press |               |               | #By default perform mouse press on (0,0)) coordinates. |
        | Mouseover Press | x_coordinates | y_coordinates | #Would perform mouse move to given coordinates. |
        """
        logger.info(" Moving mouse to given X and Y coordinates and press " + button + "button i.e: clicking at " +
                    "(" + str(x_coordinates) + "," + str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.press(button=button, coords=(x_coordinates, y_coordinates))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable move mouse to coordinates " + str(x_coordinates), str(y_coordinates) + " to perform " + \
                   button + " button click in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mouseover_release(self, x_coordinates, y_coordinates, button='left'):
        """
        Release the mouse button at specified X and Y coordinates.

        If already mose pressed on coordinates, this helps us to release the mose press.

        *Example*:
        | *Keyword*         | *Attributes*  |       |      |                                              |
        | Mouseover Release | 10  | 10  |  | #would perform mouse release on (10,10)) coordinates, if already pressed. |
        | Mouseover Release | x_coordinates | y_coordinates | right | #Would perform mouse move to given coordinates. |
        """
        logger.info(" Moving mouse to given X and Y coordinates and release " + button + "button i.e: releasing  at " +
                    "(" + str(x_coordinates) + "," + str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.release(button=button, coords=(x_coordinates, y_coordinates))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable move mouse to coordinates " + str(x_coordinates, y_coordinates) + " to perform " + button +\
                   " button release in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mousewheel_scroll(self, x_coordinates, y_coordinates, wheel_dist=1):
        """
        Do mouse wheel scroll

        If already mose pressed on coordinates, this helps us to release the mose press.

        *Example*:
        | *Keyword*         | *Attributes*  |       |      |                                              |
        | mousewheel scroll | 10  | 10  |  | #would perform mouse scroll on (10,10)) coordinates.          |
        | mousewheel scroll | x_coordinates | y_coordinates | 1 | #Would perform 1 wheel mouse scroll at coordinates. |
        """
        logger.info("scrolling mouse to given X and Y coordinates " + str(wheel_dist) +
                    " mouse wheel scroll i.e:scrolling at " + "(" + str(x_coordinates) + "," + str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.scroll(coords=(x_coordinates, y_coordinates), wheel_dist=wheel_dist)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable Do mouse wheel to coordinates " + str(x_coordinates, y_coordinates) + " to perform " + \
                   wheel_dist + " mouse wheel scroll in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def input_text(self, element_property, text):
        """
        Types the given ``text`` into the text field identified by ``element_property``.

        Type keys to the element using keyboard.send_keys.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*   | *Attributes*     |                    |
        | Input Text  | element property | XYZ                |
        | Input Text  | element property | {TAB}{SPACE}{TAB}  |

        """
        logger.info("'Typing text " + text + " into text field identified by " + element_property+"'.")
        try:
            dlg[element_property].type_keys(text)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable Types the given" + text + "into the text field identified by " + element_property + "." \
                   + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def send_keystrokes(self, text):
        """
        Silently send keystrokes to the control in an inactive window.

        It parses modifiers Shift(+), Control(^), Menu(%) and Sequences like “{TAB}”, “{ENTER}”.

        *Example*:
        | *Keyword*        | *Attributes*     |
        | Send Keystrokes  | +^               |

        """
        logger.info("sending keystrokes to the control in an inactive window.")
        try:
            dlg.send(text)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable To Silently send keystrokes to the control in an inactive window" + "." + ":::: " \
                   + "because " + str(h1)
            raise Exception(mess)

    def press_keys(self, keys):
        """
        This keyword Enter The Keys On the ``window``.

        Sends keys directly on screen into edit fields/ window.

        *Example*:
        | *Keyword*   | *Attributes*      |
        | Press Keys  | {TAB}{SPACE}{TAB} |
        | Press Keys  | ABCD              |
        """
        logger.info("'Pressing the given " + keys + " key(s) on Active Window"+"'.")
        try:
            dlg.type_keys(keys)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable To Press the given" + keys + "on Window." + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def clear_edit_field(self, element_property):
        """
        This keyword Clears if any existing content available in Given Edit Field ``element_property`` On the window.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*         | *Attributes*      |
        | Clear Edit Field  | Edit              |

        """
        logger.info("Unable Clear the given EditFields" + element_property + "on Active Window.")
        try:
            dlg[element_property].set_text(u'')
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable To Clear the given Edit Fields" + element_property + "on Active Window." + ":::: " + str(h1)
            raise Exception(mess)

    def __checkbox_status__(self, element_property, back):
        try:
            global cb_status
            if back == "uia":
                cb_status = dlg[element_property].get_toggle_state()
            else:
                print("Hima0")
                cb_status = dlg.CheckBox.is_checked()
            print(cb_status)
            return str(cb_status)
        except Exception as h1:
            mess = "couldn't find the check box status " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def select_checkbox(self, element_property, backend="uia"):
        """
        Selects(✔) the checkbox identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*         | *Attributes*      |
        | Select Checkbox   | element property  |

        """
        try:
            cb_status = self.__checkbox_status__(element_property, backend)
            logger.info("Enabling the checkbox")
            logger.info("checkbox is in unchecked state")
            logger.info("checking (✔) unchecked checkbox")
            if cb_status == str(0):
                dlg[element_property].click_input()
                logger.info("enabled(✔) unchecked checkbox successfully " + "with backend process " + backend)
            if cb_status == str(False):
                dlg.CheckBox.check()
                logger.info("enabled(✔) unchecked checkbox successfully " + "with backend process " + backend)
            elif cb_status == str(True) or str(1):
                logger.info("checkbox is already in checked state (✔)")
            # if cb_status == str(None) or str(2):
            #     logger.info("Not able to retrieve the checkbox status, un supported")
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "failed to check (✔) the check box " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def unselect_checkbox(self, element_property, backend="uia"):
        """
        Removes the selection of checkbox identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*           | *Attributes*      |
        | Unselect Checkbox   | element property  |

        """
        try:
            cb_status = self.__checkbox_status__(element_property, backend)
            logger.info("Disabling the checkbox")
            logger.info("Checkbox is in checked (✔) state")
            logger.info("Unchecking (✔-->✖) checked checkbox")
            if cb_status == str(1):
                dlg[element_property].click_input()
                logger.info("Disabled the existing checked(✔) checkbox successfully "
                            + "with backend process " + backend)
            if cb_status == str(True):
                dlg.CheckBox.uncheck()
                logger.info("Disabled the existing checked(✔) checkbox successfully " + "with backend process"
                            + backend)
            elif cb_status == str(False) or str(0):
                logger.info("checkbox is already in disabled state (✖)")
            # if cb_status == str(None) or str(2):
            #     logger.info("Not able to retrieve the checkbox status, un supported")
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "failed to uncheck the check box " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def __screenshot_on_error__(self, screenshot="Application", screenshot_format="PNG", screenshot_directory=None,
                             width="800px"):
        if self.take_screenshots is True:
            self.capture_app_screenshot(screenshot, screenshot_format, screenshot_directory, width)

    def capture_app_screenshot(self, screenshot="Application", screenshot_format="PNG", screenshot_directory=None,
                               width="800px"):
        """
        Takes a screenshot and embeds it into the log file.

        By Default image extension will be .PNG.

        This keyword supports .PNG/JPEG/JPG/TIFF/BMP image extensions.

        This keyword is used to develop to take the screenshot even in VM/Remote/Desktop,
        needs pass the screenshot name as parameter without format, default screenshot name will
        be "Application Screenshot".

        By Default all the screenshots will be located in results folder.

        Provisioned to `set screenshot_directory` in keyword as optional argument.

        *Example*:
        | *Keyword*              | *screenshot name* | *Format* | *screenshot_directory* | *width of screenshot* |
        | Capture App Screenshot |       | #By default screenshot name is Application screenshot, PNG format wth 880px |
        | Capture App Screenshot | Notepad  | TIFF |
        | Capture App Screenshot | Notepad  | JPEG | C:\\Hims | 880px |

        """
        global screenshots_directory
        if screenshot_directory is not None:
            logger.info("Taking a screenshot in " + screenshot_format + " format and keeps in folder Location " +
                        screenshot_directory + ".")
            if not os.path.exists(screenshot_directory):
                os.makedirs(screenshot_directory)
                screenshots_directory = screenshot_directory
            else:
                screenshots_directory = screenshot_directory
        else:
            logger.info("Taking a screenshot in " + screenshot_format + " format and embeds it into the log file.")
            variables = BuiltIn().get_variables()
            screenshots_directory = str(variables['${OUTPUTDIR}'])
        screenshot_format = screenshot_format.lower()
        if screenshot_format:
            try:
                index = 0
                file_path = ""
                pic = screenshot
                for i in range(1000):
                    file_path = screenshots_directory + '\\' + pic + '_' + str(i) + '.'+screenshot_format
                    filename = pic + '_' + str(i) + '.'+screenshot_format
                    file_path = re.sub('([\w\W]+)\_0(\.'+screenshot_format+')$', r'\1\2', file_path)
                    hrefsrcpath = re.sub('([\w\W]+)\_0(\.'+screenshot_format+')$', r'\1\2', filename)
                    if os.path.exists(file_path):
                        pass
                    else:
                        index = i
                        break
                if screenshot_directory is None:
                    logger.info(screenshot + '_Screenshot' + '</td></tr><tr><td colspan="3"><a href="%s"><img src="%s"\
                     width="%s"></a>' % (hrefsrcpath, hrefsrcpath, width), html=True)
                else:
                    logger.info(screenshot + '_Screenshot' + '</td></tr><tr><td colspan="3"><a href="%s"><img src="%s"\
                     width="%s"></a>' % (file_path, file_path, width), html=True)
                ph = pyautogui.screenshot()
                ph.save(file_path)
            except Exception as h1:
                mess = "Unable to take application screenshot " + ":::: " + "because " + str(h1)
                raise Exception(mess)

    def get_line(self, element_property, window_property, pattern='(.*)(L0, T0, R0, B0)(.*)', index=0, backend="uia"):
        """
        Return the line which is matching to pattern. eg:By default sub-string (L0, T0, R0, B0) containing line
        (by pattern).

        Basically when don't have object properties to perform actions or to get status, use this keyword to get
        backend properties by using matching pattern and assign boolean values to verify. Based on matching pattern
        line of property will return.

        print controls manually and use below application to make matching regex pattern.

        For patterns: follow https://regexr.com/

        refer `returning variable status`

        By Default this keyword serves with zero index application, pattern (.*)(L0, T0, R0, B0)(.*)
        and backend process of uia.

        See the `Element property` and `Window property` section for details about the locator.

        *Example*:
        | *Keyword* | *element_property* | *Window property* | *pattern* | *index* | *backend process* |
        | Get Line   | element property  | class_name:PQRS   | (.*)(L989, T878, R767, B656)(.*) | 0 | win32 |
        """
        try:
            global result
            app = pywinauto.application.Application(backend=backend)
            self._window_handle_(window_property, index)
            app.connect(handle=w_handle)
            window = app.window(handle=w_handle)
            window.set_focus()
            window[element_property].print_ctrl_ids()
            printed_controls = sys.stdout.getvalue()
            sl = printed_controls.splitlines()
            for line in sl:
                if re.match(pattern, line):
                    result = line
                    return result
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to Drawn Details: " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def returning_variable_status(self, variable, element_property, window_property, pattern,
                                  index, backend):
        """
        This keyword will return the boolean value based on variable availability in pattern matched line.

        Variable is nothing but some string. returns Status for the given variable.

        if variable available than boolean value ``Ture`` will return else ``False``.

        *Example*:
        | *Keyword* | *variable* | *element_property* | *Window property* | *pattern* | *index* | *backend process* |
        | Returning Variable Status | Some required string | element property  | class_name:ABCD   | (.*)(L989, T878, R767, B656)(.*) | 0 | win32 |

        """
        try:
            result = self.get_line(element_property, window_property, pattern, index, backend)
            if variable in result:
                return True
            else:
                return False
        except Exception as h1:
            mess = "Unable to Drawn Details: " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def close_application_window(self, window_property, backend_process="win32", index=0):
        """
        Close the window by pressing Alt+F4 keys.

        This keyword helps us to close the active application window using window handle. This keyword also helps to us
        to close application.

        Need to pass the title/class_name as arguments.

        if title available ---> title:PQR
        if class available---> class_name:XZY

        *Example*:
        | *Keyword*                 | *Attributes*      |  *backend process*  | *Index* |
        | Close Application Window  | window property   | #By default window closes where index Zero and by backend process win32 |
        | Close Application Window  | window property   | uia | 1 |
        """
        try:
            pwa_app = pywinauto.application.Application(backend=backend_process)
            self._window_handle_(window_property, index)
            pwa_app.connect(handle=w_handle)
            window = pwa_app.window(handle=w_handle)
            window.set_focus()
            window.close_alt_f4()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to close APP application" + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def __returnElementStateOnWindow__(self, element_property, window_property, index, msg=None):
        try:
            global status
            self._window_handle_(window_property, index)
            if checkup_state == 'visible':
                status = dlg[element_property].is_visible()
            if checkup_state == 'enable':
                status = dlg[element_property].is_enabled()
            return status
        except Exception as e1:
            print(e1)
            if kick_off_time >= culmination_time:
                error = AssertionError(msg) if msg else AssertionError(self)
                error.ROBOT_EXIT_ON_FAILURE = True
                raise error

    def get_element_visible_state(self, element_property):
        """
        Returns ``True`` when element visible in current window.

        returns True only when element is visible in current window.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*                  | *Attributes*       |
        | Get Element Visible State  | element property   | #would return true when element is visible. |

        """
        try:
            state = dlg[element_property].is_visible()
            return state
        except Exception as h1:
            print(h1)
            pass

    def get_element_enable_state(self, element_property):
        """
        Returns ``True`` when element enabled in current window.

        returns True only when element is visible in current window.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*                  | *Attributes*       |
        | Get Element enable State   | element property   | #would return true when element is enabled. |

        """
        try:
            state = dlg[element_property].is_enabled()
            return state
        except Exception as h1:
            print(h1)
            pass

    def wait_until_window_element_is_visible(self, element_property, window_property='title:Untitled - Notepad',
                                             index=0, timeout=60, retry_interval=1):
        """
        Waits for ``element_property`` to be made visible in given ``Window``

        Wait Until Window Element Is Visible by default with maximum timeout of 60 seconds configured.
        By default this will be for Notepad console.

        To wait for required object to visible in window, pass respective ``Element property`` and ``Window property``.
        pass as an arguments like below as shown.

        if title is available ---> title:XZY

        if class is available---> class_name:PQR

        See the `Element property` and `Window property` section for details about the locator.

        *Example*:
        | *Keyword* | *Element property* | *window property* | *app index* | *timeout* | *Retry Interval* |
        | Wait Until Window Element Is Visible | Edit  |  # This waits for Edit object to present in 60 seconds with retry interval 1sec in notepad application, |
        | Wait Until Window Element Is Visible | element property  |  class_name:xyz | 1 | 30 | 2 |
        """
        global checkup_state
        checkup_state = 'visible'
        logger.info("Waiting for element to be visible.")
        self.__finds_window_contains_status__(element_property, window_property, index, timeout, retry_interval)

    def wait_until_window_element_is_enabled(self, element_property, window_property='title:Untitled - Notepad',
                                             index=0, timeout=60, retry_interval=1):
        """
        Waits for ``element_property`` to be enabled in given ``Window``

        Wait Until Window Element Is Enabled by default with maximum timeout of 60 seconds configured.
        By default this will be for Notepad console.

        To wait for required object to enable in window, pass respective ``Element property`` and ``Window property``.
        pass as an arguments like below as shown.

        if title is available ---> title:XZY

        if class is available---> class_name:PQR

        See the `Element property` and `Window property` section for details about the locator.

        *Example*:
        | *Keyword* | *Element property* | *window property* | *app index* | *timeout* | *Retry Interval* |
        | Wait Until Window Element Is Enabled | Edit  |  # This waits for Edit object to present in 60 seconds with retry interval 1sec in notepad application, |
        | Wait Until Window Element Is Enabled | element property  |  class_name:xyz | 1 | 30 | 2 |
        """
        global checkup_state
        checkup_state = 'enable'
        logger.info("Waiting for element to be enable.")
        self.__finds_window_contains_status__(element_property, window_property, index, timeout, retry_interval)

    def __finds_window_contains_status__(self, element_property, window_property, index, timeout, retry_interval):
        try:
            global kick_off_time
            global culmination_time
            kick_off_time = datetime.now()
            culmination_time = datetime.now() + timedelta(seconds=timeout)
            while kick_off_time <= culmination_time:
                value = self.__returnElementStateOnWindow__(element_property, window_property, index)
                if value is True:
                    logger.info("Window element " + "``" + element_property + "``" + " is " + checkup_state +
                                " in Given window where " + "``" + window_property + "``")
                    return value
                else:
                    time.sleep(retry_interval)
                    kick_off_time = datetime.now()
            logger.info("Window element " + "``" + element_property + "``" + " is not " + checkup_state +
                        " in Given window where " + "``" + window_property + "``")
            raise AssertionError
        except Exception as h1:
            self.__screenshot_on_error__('timeout')
            msg = ("Timeout Error.Could not find matching element. After waiting for {0} seconds "
                   .format(str(timeout))) + str(h1)
            error = AssertionError(msg) if msg else AssertionError(self)
            error.ROBOT_EXIT_ON_FAILURE = True
            raise error

    def wait_until_window_present(self, window_property='title:Untitled - Notepad', timeout=60, retry_interval=1,
                                  index=0):
        """
        Waits for window to be present

        Wait Until window present by default with maximum timeout of 60 seconds configured. By default this will be for
        Notepad console.

        To focus any other window screen, pass respective `Window property`. Pass as an argument like below
        as shown.

        if title is available ---> title:XZY

        if class is available---> class_name:PQR

        See the `Element property` and `Window property` section for details about the locator.

        *Example*:
        | *Keyword* | *window property* | *timeout* | *Retry Interval* | *app index* |
        | Wait Until Window Present | optional Argument is WidowProperty, By Default it will look for Notepad |
        | Wait Until Window Present | title:Untitled - Notepad | #This waits for window to present in 60 seconds with retry interval 1sec in zero index notepad application. |
        | Wait Until Window Present | class_name: Save As | #By Class Name of window |
        | Wait Until Window Present | class_name: Save | 60 | 1 | 0 |
        """
        try:
            global culmination_time
            global kick_off_time
            kick_off_time = datetime.now()
            culmination_time = datetime.now() + timedelta(seconds=timeout)
            while kick_off_time <= culmination_time:
                try:
                    self._window_handle_(window_property, index)
                    logger.info("Window present with "+window_property)
                    return w_handle
                except Exception as e:
                    logger.info("Window not present with " + window_property + " "":::: " + "because " + str(e))
                    time.sleep(retry_interval)
                    kick_off_time = datetime.now()
            raise AssertionError
        except Exception as h1:
            self.__screenshot_on_error__('timeout')
            msg = ("Timeout Error.Couldn't find matching Window. After waiting for {0} seconds ".format(str(timeout)))\
                + str(h1)
            error = AssertionError(msg) if msg else AssertionError(self)
            error.ROBOT_EXIT_ON_FAILURE = True
            raise error

    def close_application(self, command):
        """
        Closes the application.

        This should only be used when it is OK to kill the process like you would do in task manager.

        Use ``task kill`` command to close the application.

        *Example*:
        | *Keyword*           | *Attributes* |
        | Close Application   | command      |

        """
        try:
            p = subprocess.Popen(command, universal_newlines=True, shell=True, stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE)
            logger.info("Running command '%s'." % command)
            output = p.stdout.read()
            output1 = p.stderr.read()
            retcode = p.wait()
            if retcode != 0:
                output = "Application is no more in active state."
                logger.info("Ret code is " + str(retcode) + " and " + output + " Verification Message: " + output1)
                return
            if 'SUCCESS:' in output:
                logger.info("Ret code is " + str(retcode) + " and Verification Message: " + output)
                self.close_application(command)
        except Exception as h1:
            logger.info(h1)

    def select_menu(self, menu_path):
        """
        Select a menu item specified in the path.

        *Example*:
        | *Keyword*     | *Attributes*   |                                 |
        | Select Menu   | Edit->Go To... | #Check notepad Edit menu option |

        """
        try:
            dlg.menu_select(menu_path)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to select the menu path " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def select_combobox_item(self, combobox, item):
        """
        Select the ComboBox item

        ``item`` argument should be given as available item from combobox list.

        ``item`` can be either a 0 based index of the item to select or it can be the string that you want to select

        *Example*:
        | *Keyword*             | *Attributes*  |            |                                                |
        | Select Combobox_Item  | ComboBox1     | item text  | #would select item in combobox                 |
        | Select Combobox_Item  | ComboBox1     | item index | #would select item in combobox based on index  |

        """
        try:
            log = ""
            if item.isdigit():
                item = int(item)
                log = "WithIndex"
            combo = dlg[combobox]
            combo.select(item)
        except Exception as h1:
            time.sleep(1)
            logger.info(h1)
            self.__screenshot_on_error__('SelectedComboItem' + log + str(item))
            pass

    def get_selected_combobox_item_index(self, combobox):
        """
        Return the selected item index.

        *Example*:
        | *Keyword*                         | *Attributes*  |                                                        |
        | Get Selected Combobox Item Index  | ComboBox2     | #would return selected item index from given combobox  |
        """
        try:
            combo = dlg[combobox]
            index = combo.selected_index()
            return index
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to select the combobox item " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_combobox_list_items(self, combobox):
        """
        Return the text of the items in the combobox

        *Example*:
        | *Keyword*                | *Attributes*  |                                                    |
        | Get Combobox List Items  | ComboBox3     | #would return items as a list from given combobox  |

        """
        try:
            combo = dlg[combobox]
            combo_list = combo.texts()
            return combo_list
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to select the combobox item " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_combobox_selected_item_text(self, combobox):
        """
        Return the selected item text.

        *Example*:
        | *Keyword*                        | *Attributes*  |                                                       |
        | Get Combobox Selected Item Text  | ComboBox4     | #would return selected item text from given combobox  |

        """
        try:
            combo = dlg[combobox]
            selected_item_text = combo.selected_text()
            return selected_item_text
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to select the combobox item " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_combobox_items_count(self, combobox):
        """
        Return the number of items in the combobox

        *Example*:
        | *Keyword*                 | *Attributes*  |                                                        |
        | Get Combobox Items count  | ComboBox5     | #would return items count present within the combobox  |

        """
        try:
            combo = dlg[combobox]
            combo_items_count = combo.item_count()
            return combo_items_count
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to select the combobox item " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def select_list_item(self, listview, item):
        """
        Select the list item

        ``item`` argument should be given as available item from list items.

        ``item`` can be either a 0 based index of the item to select or it can be the string that you want to select

        Recommended to use win32 controls. To shift to win32 controls, just use the `Focus Application Window`

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                 |
        | Focus Application Window   | title:Print |   win32 |                                            |
        | Select List Item  | ListView     | item text  | #would select item in list based on item string |
        | Select List Item  | ListView     | item index | #would select item in list based on item index  |
        """
        try:
            log=" by item string"
            if item.isdigit():
                item = int(item)
                log = "by item index"
            b = dlg[listview]
            b.set_focus()
            b.select(item)
            self.capture_app_screenshot('SelectedListItem' + log + str(item))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to select the given Item from the list view " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_list_item_text(self, listview, item):
        """
        Returns the list item text.

        ``item`` argument should be given as available item from list items.

        ``item`` can be either a 0 based index of the item to select or it can be the string that you want to select

        Recommended to use win32 controls. To shift to win32 controls, just use the `Focus Application Window`

        Recommended to use win32 controls

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                  |
        | Focus Application Window   | title:Print |  win32 |                                              |
        | Get List Item Text | ListView     | item text  | #would select item in list based on item string |
        | Get List Item Text | ListView     | item index | #would select item in list based on index       |
        """
        try:
            log = "by item string"
            if item.isdigit():
                item = int(item)
                log = "by Item Index"
            b = dlg[listview]
            list_item = b.item(item).text()
            return list_item
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "unable to get the list item text from list box " + log + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_list_items_count(self, listview):
        """
        Returns the list items count.

        Recommended to use win32 controls. To shift to win32 controls, just use the `Focus Application Window`

        *Example*:
        | *Keyword*                  | *Attributes* |                                     |
        | Focus Application Window   | title:Print  |  win32                              |
        | Get List Items Count       | ListView     | #would return items count from list |
        """
        try:
            list_items_count = dlg[listview].item_count()
            return list_items_count
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "unable to get the list items count from list box " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def check_list_item_present(self, listview, item):
        """
        Returns ``item handle`` if given item present in list, else fail.

        Recommended to use win32 controls. To shift to win32 controls, just use the `Focus Application Window`

        ``item`` can be either a 0 based index of the item to select or it can be the string that you want to select

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                      |
        | Focus Application Window   | title:Print |  win32 |                                                  |
        | Check List Item Present | ListView     | item text  | #would check item in list based on item string |
        | Check List Item Present | ListView     | item index | #would check item in list based on index       |
        """
        try:
            log = "By item string"
            if item.isdigit():
                item = int(item)
                log = "By Item Index"
            list_items = dlg[listview]
            item_handle = list_items.check(item)
            return item_handle
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "unable to check the list item presence from list box " + log + ":::: " + "because " + str(h1)
            raise Exception(mess)
