#!/usr/bin/env python

import codecs
import os
import re

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

try:
    from pypandoc import convert

    def read_md(f): return convert(f, 'rst', format='md')

except ImportError:
    print("warning: pypandoc module not found, "
          "could not convert Markdown to RST")

    def read_md(f): return open(f, 'r').read()

###################################################################

NAME = 'pcodedmp'
PACKAGES = [NAME]
META_PATH = os.path.join(NAME, NAME + '.py')
KEYWORDS = ['vba', 'p-code', 'disassembler']
CLASSIFIERS = [
    'Development Status :: 5 - Production/Stable',
    'Environment :: Console',
    'Intended Audience :: Developers',
    'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
    'Natural Language :: English',
    'Operating System :: OS Independent',
    'Programming Language :: Python',
    'Programming Language :: Python :: 2',
    'Programming Language :: Python :: 2.7',
    'Programming Language :: Python :: 3',
    'Programming Language :: Python :: 3.3',
    'Programming Language :: Python :: 3.4',
    'Programming Language :: Python :: 3.5',
    'Programming Language :: Python :: 3.6',
    'Topic :: Security',
    'Topic :: Software Development :: Disassemblers',
    'Topic :: Utilities',
]
INSTALL_REQUIRES = [
    'oletools>=0.54',
    'win_unicode_console',
]

###################################################################

HERE = os.path.abspath(os.path.dirname(__file__))


def read(*parts):
    """
    Build an absolute path from *parts* and and return the contents of the
    resulting file.  Assume UTF-8 encoding.
    """
    with codecs.open(os.path.join(HERE, *parts), 'r', 'utf-8') as f:
        return f.read()


META_FILE = read(META_PATH)


def find_meta(meta):
    """
    Extract __*meta*__ from META_FILE.
    """
    meta_match = re.search(
        r"^__{meta}__ = ['\"]([^'\"]*)['\"]".format(meta=meta),
        META_FILE, re.M
    )
    if meta_match:
        return meta_match.group(1)
    raise RuntimeError('Unable to find __{meta}__ string.'.format(meta=meta))

entry_points = {
    'console_scripts': [
        NAME + '=' + NAME + '.' + NAME + ':main',
    ],
}

if __name__ == '__main__':
    setup(
        name=NAME,
        description=find_meta('description'),
        license=find_meta('license'),
        url=find_meta('uri'),
        version=find_meta('VERSION'),
        author=find_meta('author'),
        author_email=find_meta('email'),
        maintainer=find_meta('author'),
        maintainer_email=find_meta('email'),
        keywords=KEYWORDS,
        long_description=read_md('README.md'),
        long_description_content_type='text/markdown',
        packages=PACKAGES,
        classifiers=CLASSIFIERS,
        install_requires=INSTALL_REQUIRES,
        entry_points=entry_points,
    )
