import codecs
import os

from setuptools import find_packages, setup

here = os.path.abspath(os.path.dirname(__file__))

with codecs.open(os.path.join(here, "README.md"), encoding="utf-8") as fh:
    LONG_DESCRIPTION = "\n" + fh.read()

VERSION = "1.0.4"
DESCRIPTION = "Optimised way to write in Excel files."

with open("requirements.txt") as f:
    required = f.read().splitlines()

# Setting up
setup(
    name="excel-write",
    version=VERSION,
    author="Dipan Nanda",
    author_email="d19cyber@gmail.com",
    description=DESCRIPTION,
    long_description_content_type="text/markdown",
    long_description=LONG_DESCRIPTION,
    packages=find_packages(),
    install_requires=required,
    keywords=["python", "excel", "pandas", "excel-write", "write", "openpyxl", "xlsxwriter"],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Operating System :: Unix",
        "Operating System :: MacOS :: MacOS X",
        "Operating System :: Microsoft :: Windows",
    ],
)
