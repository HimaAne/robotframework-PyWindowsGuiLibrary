from robot.libraries.BuiltIn import BuiltIn
from robot.api.deco import keyword
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

        Library----PyWindowsGuiLibrary----screenshots_on_error=False----backend=UIA

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
    __version__ = '2.1'
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'

    def __init__(self, screenshots_on_error=True, backend="uia"):
        """
        Sets default variables for the library.

        | *Arguments*                | *Documentation*                  |
        | screenshots_on_error=True  | Enables screenshots on Failures  |
        | screenshots_on_error=False | Disables screenshots on Failures |
        | backend                    |                UIA               |
        | backend                    |                win32             |

        """
        self.take_screenshots = screenshots_on_error
        self.backend = backend.lower()
        self.dlg = None
        self.APPInstance = None
        self.hae = ""

    def set_backend_win32(self):
        """
        This sets backend variable as win32 globally, for execution purpose.

        *Example*:
        | *Arguments*        | *Documentation*                 |
        | Set Backend Win32  | #sets backend process as win32. |

        """

        self.backend = 'win32'

    def set_backend_uia(self):

        """
        This sets backend variable as uia globally, for execution purpose.

        *Example*:
        | *Arguments*       | *Documentation*                 |
        | Set Backend Uia  | #sets backend process as uia.    |

        """

        self.backend = 'uia'

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

        If Required to open the Application in win32 than, need to pass backend process as win32.

        *Example*:
        | *Keyword*          | *Attributes*               |       |                                               |
        | Launch Application | C:\\Program Files\\app.exe |       | # By default application launches with uia.   |
        | Launch Application | C:\\Program Files\\app.exe | win32 | # Application launches with win32. |

        """
        logger.info("Launching the application from the path " + path)
        try:
            self.APPInstance = pywinauto.application.Application(backend=backend).start(path, timeout=60)
            return self.APPInstance
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to start/Launch given Application from the path " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def _window_handle_(self, window_property, index):
        global w_handle
        if 'title:' in window_property:
            win_prop = window_property.replace('title:', '')
            w_handle = pywinauto.findwindows.find_windows(title_re=win_prop, ctrl_index=None, top_level_only=True,
                                                          visible_only=True)[index]
        elif 'class_name:' in window_property:
            win_prop = window_property.replace('class_name:', '')
            w_handle = pywinauto.findwindows.find_windows(class_name=win_prop, ctrl_index=None, top_level_only=True,
                                                          visible_only=True)[index]

    def focus_application_window(self, window_property, backend="uia", index=0):
        """
        Focuses the Application to the Front End.

        And return the Connected Window Handle Instance :i.e dlg. This sets as global and the same used across
        further implementation.

        See the `Window property` section for details about the window.

        Use Backend argument as uia/win32 as per application support.

        This connects to the application.

        *Note:* Use this keyword anytime to shift the control of execution between backend process win32 and uia.

        *Example*:
        | *Keyword*                | *Attributes*             | *backend process* |   | |
        | Focus Application Window | title:Untitled - Notepad |       |   | # Would focus application with backend process uia and index 0.   |
        | Focus Application Window | class_name:XYZ           | uia   | 1 | # Would focus application with backend process uia and index 1.   |
        | Focus Application Window | title:ABC                | win32 | 1 | # Would focus application with backend process win32 and index 1. |

        """
        logger.info("'Focusing the application window matches to " + window_property + "'.")
        try:
            self.APPInstance = pywinauto.application.Application(backend=backend)
            self._window_handle_(window_property, index)
            self.APPInstance.connect(handle=w_handle)
            self.dlg = self.APPInstance.window(handle=w_handle)
            self.dlg.set_focus()
            if backend == "uia":
                self.set_backend_uia()
            else:
                self.set_backend_win32()
            return self.dlg
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to focus Application given by " + str(window_property) + " property,with " + str(index) \
                   + " index and " + str(backend) + " backend process " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def set_focus_to_element(self, element_property):
        """
        Sets focus to given element on current active window.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*            | *Attributes* |                                        |
        | Set Focus To Element | element_property | # Would focuses to window element. |

        """
        try:
            if self.dlg and self.backend is not None:
                self.dlg[element_property].set_focus()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to focus element in the Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_element_control_type(self, element_property):
        """
        Return the given element control type.

        See the `Element property` section for details about the locator.

        Control types are : edit, combobox, radio, button ect...

        *Example*:
        | *Variable*       | *Keyword* |  *Attributes*      |    |
        | ${ControlType}=  | Get Element Control Type | element_property | # Would return given element control type. |

        """
        try:
            if self.dlg and self.backend is not None:
                return self.dlg[element_property].friendly_class_name()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to focus element in the Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def maximize_application_window(self):
        """
        Maximize the window.

        *Example*:
        | *Keyword*                   |                                     |
        | Maximize Application Window | #would maximize application window. |

        """
        logger.info("Maximizing the current application window")
        try:
            if self.backend and self.dlg is not None:
                self.dlg.maximize()
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
        logger.info("Minimizing the current application window.")
        try:
            if self.dlg and self.backend is not None:
                self.dlg.minimize()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to minimize the Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def print_current_window_page_object_properties(self):
        """
        Returns control identifiers for current ``window``.

        Recommended to use UIA controls to get fetch more properties from backend.

        *Example*:
        | *Variable*                   | *Keyword*                                  |   |
        | ${Complete window controls}= | print_current_window_page_object_properties | #would return complete controls of window. |

        """
        logger.info("'Returns Complete window controls from Current connected Window" + "'.")
        try:
            self.dlg.print_control_identifiers()
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
        logger.info("Returns Specific element controls from given ``element Property`` " + "``" + element_property
                    + "`` " + "'.")
        try:
            self.dlg[element_property].print_ctrl_ids()
            printed_controls = sys.stdout.getvalue()
            return printed_controls
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to print and return Object " + element_property + " properties" + ":::: " + "because " \
                   + str(h1)
            raise Exception(mess)

    def get_window_text(self):
        """
        Returns the text of the currently selected window.

         *Example*:
        | *Variable*      |  *Keyword*       |                             |
        | ${Window Text}= | Get Window Text  | #would returns window text. |

        """
        logger.info("Retrieving current window text.")
        try:
            if self.backend is not None:
                return self.dlg.window_text()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to get current window text in current active Application " + " :::: " + "because " \
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
        logger.info("Retrieving element text from given Object Property" + "'" + element_property + "'.")
        try:
            if self.dlg and self.backend is not None:
                return self.dlg[element_property].window_text()
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
        logger.info("Setting text " + text + " into text field identified by " + "'" + element_property + "'.")
        try:
            if self.dlg and self.backend is not None:
                self.dlg[element_property].set_text(text)           # set_edit_text(text)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to Type the given " + text + " into the text field identified by " + element_property + "." \
                   + " :::: " + "because " + str(h1)
            raise Exception(mess)

    def click_on_element(self, element_property):
        """
        Clicks at the element identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*        | *Attributes*     |                                               |
        | Click On Element | element property | # Would click element with by Click() method. |

        """
        logger.info("Clicking element Property " + "'" + element_property + "'.")
        try:
            if self.backend is not None:
                self.dlg[element_property].click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click element " + element_property + " in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def double_click_on_element(self, element_property):
        """
        Double Click at the element identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*               | *Attributes*     |                                                      |
        | Double Click On Element | element property | # Would click element with by Double_Click() method. |

        """
        logger.info("Double clicking element property " + "`" + element_property + "`.")
        try:
            if self.dlg and self.backend is not None:
                if self.backend == "win32":
                    self.dlg[element_property].double_click()
                elif self.backend == "uia":
                    logger.info("This keyword is no longer support for uia controls")
                    raise AssertionError
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click object " + element_property + " in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def right_click_on_element(self, element_property):
        """
        Right Click at the element identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*               | *Attributes*     |                                                      |
        | Right Click On Element | element property  | # Would click element with by Right_Click() method.  |

        """
        logger.info("Right clicking element property " + "'" + element_property + "'.")
        try:
            if self.dlg and self.backend is not None:
                if self.backend == "win32":
                    self.dlg[element_property].right_click()
                elif self.backend == "uia":
                    log = "This keyword is no longer support for uia controls, Use other Click related keywords"
                    logger.info(log)
                    raise AssertionError
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click object " + element_property + " in Application " + ":::: " + "because " + \
                   log + str(h1)
            raise Exception(mess)

    def action_right_click_on_element(self, element_property):
        """
        Realistic Right Click at the element identified by ``element_property``.

        This is different from click method in that it requires the control to be visible on the screen but performs
        a more realistic ‘click’ simulation.

        This method is also vulnerable if the mouse is moved by the user as that could easily move the mouse off the
        control before the click_input has finished.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*                     | *Attributes*     |                                                      |
        | Action Right Click On Element | element property | # Would click element with by right_Click_input() method. |

        """
        logger.info("Action Right Clicking element Property " + "'" + element_property + "'.")
        try:
            if self.dlg and self.backend is not None:
                if self.backend == "win32":
                    self.dlg[element_property].right_click_input()
                elif self.backend == "uia":
                    log = "This keyword is no longer support for uia controls, Use other Click related keywords"
                    logger.info(log)
                    raise AssertionError
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click object " + element_property + " in Application " + ":::: " + "because " + log + \
                   str(h1)
            raise Exception(mess)

    def action_click(self, element_property):
        """
        Clicks at the element identified by ``element_property``.

        Use Element property or child window to perform the clicking operation.

        This is different from click method in that it requires the control to be visible on the screen but performs
        a more realistic ‘click’ simulation.

        This method is also vulnerable if the mouse is moved by the user as that could easily move the mouse off the
        control before the click_input has finished.

        See the `Element property` and `Child window` section for details about the locator.

        *Example*:
        | *Keyword*        | *Attributes*     |                                                 |
        | Action Click | element property | # Would click element with by Click_input() method. |
        | Action Click | child window     | # Would click element with by Click_input() method. |

        """
        logger.info("Realistic clicking element property " + "'" + element_property + "'.")
        try:
            if self.dlg and self.backend is not None:
                self.dlg[element_property].click_input()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable to click object " + element_property + " in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def action_double_click(self, element_property):
        """
        Double Click at the element identified by ``element_property``.

        Use Element property or child window to perform the clicking operation.

        This is different from click method in that it requires the control to be visible on the screen but performs
        a more realistic ‘click’ simulation.

        This method is also vulnerable if the mouse is moved by the user as that could easily move the mouse off the
        control before the click_input has finished.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*           | *Attributes*     |                                                            |
        | Action Double Click | element property | # Would click element with by Double_Click_input() method. |
        | Action Double Click | child window     | # Would click element with by Double_Click_input() method. |

        """
        logger.info("Realistic double clicking element Property " + "'" + element_property + "'.")
        try:
            if self.dlg and self.backend is not None:
                self.dlg[element_property].double_click_input()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = " Unable to perform realistic double click on object " + element_property + " in Application " \
                   + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mouseover_click(self, x_coordinates='0', y_coordinates='0', button='left'):
        """
        Clicks at the specified X and Y coordinates.

        By Default this performs left mouse click

        *Example*:
        | *Keyword*       | *Attributes*  |               |       |                                                  |
        | Mouseover Click |      |      |                    | #By default perform left click on (0,0)  coordinates. |
        | Mouseover Click | x_coordinates | y_coordinates | right | #Would perform right click on given coordinates. |
        | Mouseover Click | x_coordinates | y_coordinates | left | #Would perform left click on given coordinates.   |

        """
        logger.info(button + " Mouse button clicking at X and Y coordinates i.e " + "(" + str(x_coordinates) + "," +
                    str(y_coordinates) + ")")
        try:
            x_coordinates = int(x_coordinates)
            y_coordinates = int(y_coordinates)
            pywinauto.mouse.click(button=button, coords=(x_coordinates, y_coordinates))
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click object by " + str(x_coordinates), str(y_coordinates) + " in Application " + \
                   ":::: " + "because " + str(h1)
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
            mess = " Unable to double click object by " + str(x_coordinates), str(y_coordinates) + " in Application " \
                   + ":::: " + "because " + str(h1)
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
            mess = " Unable move mouse to coordinates " + str(x_coordinates), str(y_coordinates) + " in Application "\
                   + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def mouseover_press(self, x_coordinates='0', y_coordinates='0', button='left'):
        """
        Press the mouse button at specified X and Y coordinates.

        *Example*:
        | *Keyword*       | *Attributes*  |               |                                                 |
        | Mouseover Press |               |               | #By default perform mouse press on (0,0)) coordinates. |
        | Mouseover Press | x_coordinates | y_coordinates | #Would perform mouse press to given coordinates. |
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

        If already mose pressed on coordinates, this helps us to release the mouse press.

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
                   str(wheel_dist) + " mouse wheel scroll in Application " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def input_text(self, element_property, text):
        """
        Types the given ``text`` into the text field identified by ``element_property``.

        Type keys to the element using keyboard.send_keys.

        It parses modifiers Shift(+), Control(^), Menu(%) and Sequences like “{TAB}”, “{ENTER}”.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*   | *Attributes*     |                    |
        | Input Text  | element property | XYZ                |
        | Input Text  | element property | {TAB}{SPACE}{TAB}  |

        """
        logger.info("Typing text " + text + " into text field identified by " + element_property + ".")
        try:
            if self.dlg and self.backend is not None:
                self.dlg[element_property].type_keys(text, pause=0.05, with_spaces=True, with_tabs=True,
                                                     with_newlines=True, turn_off_numlock=True)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to types the given" + text + "into the text field identified by " + element_property + "." \
                   + " :::: " + "because " + str(h1)
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
            if self.dlg and self.backend is not None:
                if self.backend == 'win32':
                    self.dlg.send_keystrokes(text)
                elif self.backend == "uia":
                    log = "This keyword is no longer support for uia controls, Use other send/type related keywords"
                    logger.info(log)
                    raise AssertionError
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to silently send keystrokes to the control in an inactive window. " + " :::: " \
                   + "because " + log + str(h1)
            raise Exception(mess)

    def press_keys(self, keys, press_count=1, interval=0):
        """
        This keyword Enters/performs the Keys actions On the ``window``.

        Sends keys directly on screen into edit fields/ window by pyautogui.

        The following are the valid strings to pass to the press() function:
        https://pyautogui.readthedocs.io/en/latest/keyboard.html?highlight=press#keyboard-keys

        press count helps to press the same key multiple times

        *Example*:
        | *Keyword*   |  *Attributes*      |     |    |    |
        | Press Keys  |  tab               |  2  | #would click tab 2 times |    |
        | Press Keys  |  A          |  5  | 1 | #would click 'A' 5 times and for each time with 1 second time interval |
        | Press Keys  |  enter             | #would press the enter button |   |   |

        """
        logger.info("'Pressing the given " + keys + " key(s) on Active Window"+"'.")
        try:
            time.sleep(0.25)
            pyautogui.press(keys, presses=press_count, interval=interval)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to press the given" + keys + "on Window." + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def text_writer(self, string, interval=0):
        """
        The primary keyboard function is that types given string in to where cursor present at edit fields.

        This function will type the characters in the string that is passed at cursor point. To add a delay interval
        in between pressing each character key, pass an int or float for the interval keyword argument.

        types given string directly on screen into edit fields/ window by pyautogui where cursor focused.

        The following are the valid strings to pass to the write() function:
        https://pyautogui.readthedocs.io/en/latest/keyboard.html?highlight=write#the-write-function

        *Example*:
        | *Keyword*    |  *Attributes*  |      |    |
        | Text Writer  |  HimaAne  | #would type the given string with out any delay between character |   |
        | Text Writer  |  ABCD | 0.25 | #would type the given string with 0.25delay between each character |

        """
        logger.info("'Pressing the given " + "`" + string + "`" + " string on Active Window" + "'.")
        try:
            time.sleep(0.25)
            pyautogui.write(string, interval=interval)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to write/type the given " + "`" + string + "`" + " on Window where cursor placed." +\
                   ":::: " + "because " + str(h1)
            raise Exception(mess)

    def clear_edit_field(self, element_property):
        """
        This keyword Clears if any existing content available in Given Edit Field ``element_property`` On the window.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*         | *Attributes*      |
        | Clear Edit Field  | Edit              |

        """
        logger.info("clearing the given edit fields " + element_property + " on Active Window.")
        try:
            if self.dlg and self.backend is not None:
                self.dlg[element_property].set_text(u'')
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to clear the given edit field `" + element_property + "` on Active Window." + " :::: "\
                   + str(h1)
            raise Exception(mess)

    def __checkbox_status__(self, element_property):
        try:
            cb_status = ""
            if self.backend == "uia" and self.dlg is not None:
                cb_status = self.dlg[element_property].get_toggle_state()
            elif self.backend == "win32" and self.dlg is not None:
                cb_status = self.dlg.CheckBox.is_checked()
            return str(cb_status)
        except Exception as h1:
            mess = "couldn't find the check box status " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def select_checkbox(self, element_property):
        """
        Selects(✔) the checkbox identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*         | *Attributes*      |
        | Select Checkbox   | element property  |

        """
        try:
            if self.dlg and self.backend is not None:
                cb_status = self.__checkbox_status__(element_property)
                logger.info("Enabling the checkbox")
                logger.info("Checkbox is in unchecked state")
                logger.info("checking (✔) unchecked checkbox")
                if cb_status == str(0) or cb_status == str(False):
                    self.dlg[element_property].click_input()
                    logger.info("Changed checkbox state from unchecked to checked(✔) successfully. ")
                elif cb_status == str(True) or cb_status == str(1):
                    logger.info("Checkbox is already in enabled/checked state (✔)")
                else:
                    logger.info("Not able to retrieve the checkbox status, un supported & checkbox status is " +
                                cb_status)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to check (✔) the check box " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def unselect_checkbox(self, element_property):
        """
        Removes the selection of checkbox identified by ``element_property``.

        See the `Element property` section for details about the locator.

        *Example*:
        | *Keyword*           | *Attributes*      |
        | Unselect Checkbox   | element property  |

        """
        try:
            if self.dlg and self.backend is not None:
                cb_status = self.__checkbox_status__(element_property)
                logger.info("Disabling the checkbox")
                logger.info("Checkbox is in checked (✔) state")
                logger.info("Unchecking (✔-->✖) checked checkbox")
                if cb_status == str(1) or cb_status == str(True):
                    self.dlg[element_property].click_input()
                    logger.info("Changed checkbox state from checked(✔) to unchecked successfully. ")
                elif cb_status == str(False) or cb_status == str(0):
                    logger.info("Checkbox is already in disabled/Unchecked state (✖)")
                else:
                    logger.info("Not able to retrieve the checkbox status, un supported & checkbox status is "
                                + cb_status)
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Failed to uncheck the check box " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def __screenshot_on_error__(self, screenshot="Application", screenshot_format="PNG", screenshot_directory=None,
                                width="800px"):
        if self.take_screenshots is True:
            self.capture_app_screenshot(screenshot, screenshot_format, screenshot_directory, width)

    @keyword
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
                mess = "Unable to take application screenshot at " + str(index) + ":::: " + "because " + str(h1)
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
            if element_property is not None:
                window[element_property].print_ctrl_ids()
            elif element_property is None:
                window.print_ctrl_ids()
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
            mess = "Failed to close application window" + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def __returnElementStateOnWindow__(self, element_property, window_property, index, msg=None):
        try:
            global status
            self._window_handle_(window_property, index)
            if checkup_state == 'visible':
                status = self.dlg[element_property].is_visible()
            if checkup_state == 'enable':
                status = self.dlg[element_property].is_enabled()
            return status
        except Exception as e1:
            print(e1)
            if kick_off_time >= culmination_time:
                error = AssertionError(msg) if msg else AssertionError(self)
                error.ROBOT_EXIT_ON_FAILURE = True
                raise error

    @keyword
    def get_element_visible_state(self, element_property):
        """
        Returns ``True`` when element visible in current window.

        returns True only when element is visible in current window.

        See the `Element property` section for details about the locator.

        Make sure your application response while using this keyword.ss

        *Example*:
        | *Keyword*                  | *Attributes*       |                                                  |
        | Get Element Visible State  | element property   | #would return only true when element is visible. |

        """
        try:
            if self.dlg and self.backend is not None:
                state = self.dlg[element_property].is_visible()
                return state
        except Exception as h1:
            print(h1)
            pass

    @keyword
    def get_element_enable_state(self, element_property):
        """
        Returns ``True`` when element enabled in current window.

        returns True only when element is visible in current window.

        See the `Element property` section for details about the locator.

        Make sure your application response while using this keyword.

        *Example*:
        | *Keyword*                  | *Attributes*       |                                             |
        | Get Element enable State   | element property   | #would return true when element is enabled. |

        """
        try:
            if self.dlg and self.backend is not None:
                state = self.dlg[element_property].is_enabled()
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
            ret_code = p.wait()
            if ret_code != 0:
                output = "Application is no more in active state."
                logger.info("Ret code is " + str(ret_code) + " and " + output + " Verification Message: " + output1)
                return
            if 'SUCCESS:' in output:
                logger.info("Ret code is " + str(ret_code) + " and Verification Message: " + output)
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
            if self.dlg and self.backend is not None:
                self.dlg.menu_select(menu_path)
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
            if self.dlg and self.backend is not None:
                log = ""
                if item.isdigit():
                    item = int(item)
                    log = "WithIndex"
                combo = self.dlg[combobox]
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
            if self.dlg and self.backend is not None:
                combo = self.dlg[combobox]
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
            if self.dlg and self.backend is not None:
                combo = self.dlg[combobox]
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
            if self.dlg and self.backend is not None:
                combo = self.dlg[combobox]
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
            if self.dlg and self.backend is not None:
                combo = self.dlg[combobox]
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

        To shift to win32 controls, just use the `Focus Application Window`

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                 |
        | Focus Application Window   | title:Print |   win32 |                                            |
        | Select List Item  | ListView     | item text  | #would select item in list based on item string |
        | Select List Item  | ListView     | item index | #would select item in list based on item index  |

        """
        try:
            if self.dlg and self.backend is not None:
                log = " by item string"
                if item.isdigit():
                    item = int(item)
                    log = "by item index"
                b = self.dlg[listview]
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

        Use `Focus Application Window` keyword anytime to shift the control of execution between backend process
        win32 and uia.

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                  |
        | Focus Application Window   | title:Print |  win32 |                                              |
        | Get List Item Text | ListView     | item text  | #would select item in list based on item string |
        | Get List Item Text | ListView     | item index | #would select item in list based on index       |

        """
        try:
            if self.dlg and self.backend is not None:
                log = "by item string"
                if item.isdigit():
                    item = int(item)
                    log = "by Item Index"
                b = self.dlg[listview]
                list_item = b.item(item).text()
                return list_item
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to get the list item text from list box " + log + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_list_items_count(self, listview):
        """
        Returns the list items count.

        Use `Focus Application Window` keyword anytime to shift the control of execution between backend process
        win32 and uia.

        *Example*:
        | *Keyword*                  | *Attributes* |                                     |
        | Focus Application Window   | title:Print  |  win32                              |
        | Get List Items Count       | ListView     | #would return items count from list |

        """
        try:
            if self.dlg and self.backend is not None:
                list_items_count = self.dlg[listview].item_count()
                return list_items_count
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to get the list items count from list box " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def check_list_item_present(self, listview, item):
        """
        Returns ``item handle`` if given item present in list, else return failure.

        Use `Focus Application Window` keyword anytime to shift the control of execution between backend process
        win32 and uia.

        ``item`` can be either a 0 based index of the item to select or it can be the string that you want to select

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                      |
        | Focus Application Window   | title:Print |  win32 |                                                  |
        | Check List Item Present | ListView     | item text  | #would check item in list based on item string |
        | Check List Item Present | ListView     | item index | #would check item in list based on index       |

        """
        try:
            if self.dlg and self.backend is not None:
                log = "By item string"
                if item.isdigit():
                    item = int(item)
                    log = "By Item Index"
                list_items = self.dlg[listview]
                item_handle = list_items.check(item)
                return item_handle
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to check the list item presence from list box " + log + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def get_systemTray_icon_Text(self, app_child_window_property, timeout=1,
                                 taskbar_win_property='class_name:Shell_TrayWnd',
                                 show_hidden_icons_property="title:Notification Chevron",
                                 sys_tray_property="class_name:NotifyIconOverflowWindow", backend='uia'):
        """
        Returns the text of systemTray Icon.

        `timeout`, `taskbar_win_property` , `show_hidden_icons_property` , `sys_tray_property` , `backend`
        is optional arguments. If property changes use appropriate properties.

        Increase and use timeout argument for sync purpose to perform the get text of application action.

        This keyword internally handles: connecting to taskbar, clicking on show hidden icon (^) from task bar,
        takes application control from tray icons to perform the actions.

        *Example*:
        | *Keyword*         | *Attributes* |            |                                                            |
        | Get SystemTray Icon Text | title:notepad |  | #would return app text with by default 1 second time wait    |
        | Get SystemTray Icon Text | class_name:note | 2 | #would return app text from systemTray with 2seconds wait |

        """
        logger.info("Getting text of " + "`" + app_child_window_property + "`" + " application from system tray")
        try:
            self.__openSysTray_ConnectTray_TakeAppChildWindowControl__(app_child_window_property, timeout,
                                                                       taskbar_win_property, show_hidden_icons_property,
                                                                       sys_tray_property, backend)
            iconText = self.hae.window_text()
            logger.info(iconText)
            return iconText
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to get the text of application Icon which is present in the systemTray " + ":::: " + \
                   "because " + str(h1)
            raise Exception(mess)

    def click_on_show_hidden_icons(self, taskbar_win_property, show_hidden_icons_property, backend):
        """
        clicks on show hidden icons (^) from task bar.

        *Example*:
        | *Keyword*            | *Attributes-->*  *taskbar_win_property*  |  *show_hidden_icons_property* | *backend* |
        | Click On Show Hidden Icons | class_name:Shell_TrayWnd  | title:Notification Chevron    | uia |

        """
        logger.info("Clicking at show hidden icons (^) option from task bar")
        try:
            self.connect_taskbar(taskbar_win_property, backend)
            self.__apps_child_window_handle__(show_hidden_icons_property)
            self.hae.click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to open show hidden icons (From Tool Bar-->^) " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def __connect_to_tray_handle__(self, timeout, sys_tray_property, backend):
        self.connect_app(sys_tray_property, backend)
        self.hae = self.dlg
        self.hae.wait('visible', timeout=timeout, retry_interval=.25)

    def connect_app(self, window_property, backend):
        """
        connects to application window, it will only perform connect function based on title/class_name of the
        applications.

        it will made connection to application and returns reference in self.dlg variable.

        *Example*:
        | *Keyword*            | *Attributes-->* *window_property* | *backend* |
        | Connect App          | class_name:Shell_TrayWnd          | uia       |
        | Connect App          | title:Shell_TrayWnd               | win32     |

        """
        if 'title:' in window_property:
            win_prop = window_property.replace('title:', '')
            app = pywinauto.application.Application(backend=backend).connect(title=win_prop)
            self.dlg = app.window(title=win_prop)
        elif 'class_name:' in window_property:
            win_prop = window_property.replace('class_name:', '')
            app = pywinauto.application.Application(backend=backend).connect(class_name=win_prop)
            self.dlg = app.window(class_name=win_prop)
    
    def connect_taskbar(self, taskbar_win_property, backend):
        """
        connects to window taskbar, it will made connection to taskbar and returns reference in self.hae variable.

        Use Shell_TrayWnd class name to connect taskbar

        *Example*:
        |  *Keyword*                | *Attributes-->* *window_property* | *backend* |
        |  Connect Taskbar          | class_name:Shell_TrayWnd          | uia       |

        """
        logger.info("Connecting to task bar")
        self.connect_app(taskbar_win_property, backend)
        self.hae = self.dlg

    def click_on_systemTray_icon(self, app_child_window_property, timeout=1,
                                 taskbar_win_property='class_name:Shell_TrayWnd',
                                 show_hidden_icons_property="title:Notification Chevron",
                                 sys_tray_property='class_name:NotifyIconOverflowWindow', backend='uia'):
        """
        Performs Click on application Icon presents in SystemTray.

        `timeout`, `taskbar_win_property` , `show_hidden_icons_property` , `sys_tray_property` , `backend`
        is optional arguments. If property changes use appropriate properties.

        Increase and use timeout argument for sync purpose to perform the click action.

        This keyword internally handles: connecting to taskbar, clicking on show hidden icon (^) from task bar,
        takes application control from tray icons to perform the click action on tray apps.

        *Example*:
        | *Keyword*                | *Attributes* |
        | Click On SystemTray Icon | title:Microsoft Teams - New activity |

        """
        logger.info("Clicking on " + "`" + app_child_window_property + "`" + " application from system tray")
        try:
            self.__openSysTray_ConnectTray_TakeAppChildWindowControl__(app_child_window_property, timeout,
                                                                       taskbar_win_property, show_hidden_icons_property,
                                                                       sys_tray_property, backend)
            self.hae.click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to click on application icon which is presents in systemTray " + ":::: " + "because " + \
                   str(h1)
            raise Exception(mess)

    def double_click_on_systemTray_icon(self, app_child_window_property, timeout=1,
                                        taskbar_win_property='class_name:Shell_TrayWnd',
                                        show_hidden_icons_property="title:Notification Chevron",
                                        sys_tray_property="class_name:NotifyIconOverflowWindow", backend='uia'):
        """
        Performs Double clicks on application Icon presents in SystemTray.

        `timeout`, `taskbar_win_property` , `show_hidden_icons_property` , `sys_tray_property` , `backend`
        is optional arguments. If property changes use appropriate properties.

        Increase and use timeout argument for sync purpose to perform the click action.

        This keyword internally handles: connecting to taskbar, clicking on show hidden icon (^) from task bar,
        takes application control from tray icons to perform the double click action on tray apps.

        *Example*:
        | *Keyword*                       | *Attributes*                         |   |
        | Double Click On SystemTray Icon | title:Microsoft Teams - New activity |   |
        | Double Click On SystemTray Icon | title:Microsoft Teams - New activity | 2 |

        """
        logger.info("Double clicking on " + "`" + app_child_window_property + "`" + " application from system tray")
        try:
            self.__openSysTray_ConnectTray_TakeAppChildWindowControl__(app_child_window_property, timeout,
                                                                       taskbar_win_property, show_hidden_icons_property,
                                                                       sys_tray_property, backend)
            self.hae.double_click_input()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to double click on application icon which is presents in systemTray " + ":::: " + \
                   "because " + str(h1)
            raise Exception(mess)

    def __apps_child_window_handle__(self, window_property):
        if 'title:' in window_property:
            win_prop = window_property.replace('title:', '')
            self.hae = self.hae.child_window(title=win_prop).wrapper_object()
        elif 'class_name:' in window_property:
            win_prop = window_property.replace('class_name:', '')
            self.hae = self.hae.child_window(class_name=win_prop).wrapper_object()

    def right_click_on_systemTray_icon(self, app_child_window_property, timeout=1,
                                       taskbar_win_property='class_name:Shell_TrayWnd',
                                       show_hidden_icons_property="title:Notification Chevron",
                                       sys_tray_property="class_name:NotifyIconOverflowWindow", backend='uia'):
        """
        Performs Right click on application Icon presents in SystemTray.

        `timeout`, `taskbar_win_property` , `show_hidden_icons_property` , `sys_tray_property` , `backend`
        is optional arguments. If property changes use appropriate properties.

        Increase and use timeout argument for sync purpose to perform the click action.

        This keyword internally handles: connecting to taskbar, clicking on show hidden icon (^) from task bar,
        takes application control from tray icons to perform the right click action on tray apps.

        *Example*:
        | *Keyword*                      | *Attributes* |
        | Right Click On SystemTray Icon | title:Microsoft Teams - New activity |

        """
        logger.info("right clicking on " + "`" + app_child_window_property + "`" + " application from system tray")
        try:
            self.__openSysTray_ConnectTray_TakeAppChildWindowControl__(app_child_window_property, timeout,
                                                                       taskbar_win_property, show_hidden_icons_property,
                                                                       sys_tray_property, backend)
            self.hae.click_input(button="right")
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to right click on application icon which is present in systemTray Icon " + ":::: " + \
                   "because " + str(h1)
            raise Exception(mess)

    def __openSysTray_ConnectTray_TakeAppChildWindowControl__(self, app_child_window_property, timeout,
                                                              taskbar_win_property, show_hidden_icons_property,
                                                              sys_tray_property, backend):
        self.click_on_show_hidden_icons(taskbar_win_property, show_hidden_icons_property, backend)
        time.sleep(timeout)
        logger.info("Connecting to system tray handle")
        self.__connect_to_tray_handle__(timeout, sys_tray_property, backend)
        logger.info("Taking control of app child window present in system tray to perform actions")
        self.__apps_child_window_handle__(app_child_window_property)

    def click_on_window_start(self, windows_icon_property='class_name:Start',
                              taskbar_win_property='class_name:Shell_TrayWnd', backend='uia'):
        """
        Clicks at windows Icon at task bar.

        `windows_icon_property` , `taskbar_win_property` and `backend` are optional arguments. If property changes use
        appropriate properties.

        Clicking based on windows class property by default.

        *Example*:
        | *Keyword*             | *Attributes*                    |
        | Click On Window Start | #would clicks on windows button |

        """
        try:
            self.connect_taskbar(taskbar_win_property, backend)
            self.__apps_child_window_handle__(windows_icon_property)
            logger.info("Clicking on Windows Icon at taskbar")
            self.hae.click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to clicks on windows start button from task bar " + ":::: " + "because " + str(h1)
            raise Exception(mess)

    def click_at_windows_type_here_to_search_option(self, button='Button1',
                                                    taskbar_win_property='class_name:Shell_TrayWnd', backend='uia'):
        """
        Clicks at window's "type here to search" in task bar.

        `button` , `taskbar_win_property` and `backend` are optional arguments. If property changes use
        appropriate properties.

        *Example*:
        | *Keyword*                                   | *Attributes*   |                       |
        | Click At Windows Type here To Search Option | #would clicks on windows Search option |          |
        | Click At Windows Type here To Search Option |  Button1 | #would clicks on windows Search option |

        """
        try:
            self.connect_taskbar(taskbar_win_property, backend)
            logger.info("Clicking on ``Type here to search`` in task bar.")
            self.hae[button].click()
        except Exception as h1:
            self.__screenshot_on_error__()
            mess = "Unable to clicks on windows type here search option button from task bar " + ":::: " + \
                   "because " + str(h1)
            raise Exception(mess)
