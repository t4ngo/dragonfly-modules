#
# Script to generate documentation for Dragonfly command-modules.
#

import sys
import os
import os.path
import shutil
import subprocess
import textwrap
from glob import glob

directory = os.path.dirname(__file__)


#---------------------------------------------------------------------------
# Formatter class to help generate command-module documentation.

class Formatter(object):

    def __init__(self):
        self.lines = []

    def add_section(self, title):
        self.lines.extend([title, "="*len(title), ""])

    def add_subsection(self, title):
        self.lines.extend([title, "-"*len(title), ""])

    def add_subsubsection(self, title):
        self.lines.extend([title, "^"*len(title), ""])

    def add_definition(self, name, text):
        text = textwrap.wrap(text, initial_indent="   ", subsequent_indent="   ")
        self.lines.extend([name, text, ""])

    def add_textblock(self, text):
        self.lines.extend(text.splitlines())
        self.lines.append("")

    def add_codeblock(self, text):
        self.lines.extend([
            ".. code-block:: python",
            "   :linenos:",
            "",
            ])
        prefix = "   "
        lines = (prefix + l for l in text.splitlines())
        self.lines.extend(lines)
        self.lines.append("")


#---------------------------------------------------------------------------
# Extract docscrings from command-modules.

def extract_docstring(src_path, dst_path):
    namespace = {"__file__": src_path}
    try:
        execfile(src_path, namespace)
    except Exception, e:
        raise

    formatter = Formatter()

    if "__doc__" in namespace:
        module_doc = namespace["__doc__"]
        module_doc = textwrap.dedent(module_doc)
        formatter.add_textblock(module_doc)

    src_base = os.path.splitext(src_path)[0]
    config_pattern = "%s*.txt" % src_base
    config_files = glob(config_pattern)
    if config_files:
        formatter.add_subsection("Configuration examples")
        for path in config_files:
            basename = os.path.basename(path)
            print "processing configuration %r" % basename
            formatter.add_subsubsection('Configuration "%s"' % basename)
            content = open(path).read()
            formatter.add_codeblock(content)

    formatter.add_subsection("Module source code")
    location = "http://dragonfly-modules.googlecode.com/svn/trunk/command-modules/%s" % os.path.basename(src_path)
    download_text = "The most current version of this module can be" \
                    " downloaded from the `repository here <%s>`_." \
                    % location
    formatter.add_textblock(download_text)
    formatter.add_codeblock(open(src_path).read())

    f = open(dst_path, "w")
    output = "\n".join(formatter.lines)
    f.write(output)


#---------------------------------------------------------------------------
# Generate command-module documentation.

mod_dir = os.path.abspath(os.path.join(directory, "command-modules"))
for filename in os.listdir(mod_dir):
    base, extension = os.path.splitext(filename)
    if extension != ".py": continue
    print "command module:", filename

    src_path = os.path.join(mod_dir, filename)
    dst_path = os.path.join(directory, "documentation", "mod-%s.txt" % base)
    extract_docstring(src_path, dst_path)


#---------------------------------------------------------------------------
# Generate documentation.

python_binary = sys.executable
build_binary = r"c:\python25\scripts\sphinx-build-script.py"
build_type = "html"
src_dir = os.path.abspath(os.path.join(directory, "documentation"))
dst_dir = os.path.abspath(os.path.join(directory, "command-modules", "documentation"))

arguments = [python_binary, build_binary, "-a", "-b", build_type, src_dir, dst_dir]
print "Executing:", arguments
subprocess.call(arguments)
