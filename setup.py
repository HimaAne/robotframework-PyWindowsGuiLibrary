from setuptools import setup, find_packages
from os import path
this_directory = path.abspath(path.dirname(__file__))
with open('README.md') as f:
    long_description = f.read()

setup(
    name="robotframework-PyWindowsGuiLibrary",
    version="1.1",
    author="Himavanthudu Bodapati",
    author_email="himavanthudu.b@gmail.com",
    description="A Robot Framework Library for automating the WINDOWS BASED GUI applications",
    long_description=long_description,
    long_description_content_type='text/markdown',
    url="https://github.com/HimaAne/robotframework-PyWindowsGuiLibrary/archive/refs/tags/1.2.tar.gz",
    packages=find_packages(),
    classifiers=(
        "Programming Language :: Python :: 3.6",
        "Operating System :: Microsoft :: Windows",
    ),
    install_requires=["pywinauto>=0.6.8", "robotframework>=4.0.3", "pyautogui>=0.9.53", "datetime>=4.3"]
)
