from setuptools import setup, find_packages
setup(
    name = "msodumper",
    version = "0.3.0",
    packages = find_packages(),
    scripts = ['ppt-dump.py'],
    zip_safe = True,
    author = "Kohei Yoshida <kyoshida@novell.com>, Thorsten Behrens <tbehrens@novell.com>",
    author_email = "libreoffice@lists.freedesktop.org",
    description = "A package for analysing and dumping MS office formats",
    license = "MPL/LGPLv3+",
    url = "http://cgit.freedesktop.org/libreoffice/contrib/mso-dumper/"
)
