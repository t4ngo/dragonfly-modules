#
# Script to generate documentation for Dragonfly command-modules.
#

import sys
import os
import os.path
import shutil
import subprocess
import textwrap

directory = os.path.dirname(__file__)


class Formatter(object):

    def __init__(self):
        self.lines = []

    def add_section(self, title):
        self.lines.extend([title, "="*len(title), ""])

    def add_subsection(self, title):
        self.lines.extend([title, "-"*len(title), ""])

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


# Extract docscrings from command-modules.
def extract_docstring(src_path, dst_path):
    namespace = {"__file__": src_path}
    try:
        execfile(src_path, namespace)
    except Exception, e:
        raise

    formatter = Formatter()

    basename = os.path.basename(src_path)
    formatter.add_section("Command-module %s" % basename)

    if "__doc__" in namespace:
        module_doc = namespace["__doc__"]
        module_doc = textwrap.dedent(module_doc)
        formatter.add_textblock(module_doc)

    formatter.add_subsection("Module source code")
    formatter.add_codeblock(open(src_path).read())

    f = open(dst_path, "w")
    output = "\n".join(formatter.lines)
    f.write(output)


mod_dir = os.path.abspath(os.path.join(directory, "command-modules"))
for filename in os.listdir(mod_dir):
    base, extension = os.path.splitext(filename)
    if extension != ".py": continue
    print "command module:", filename

    src_path = os.path.join(mod_dir, filename)
    dst_path = os.path.join(directory, "doc-src", "mod-%s.txt" % base)
    extract_docstring(src_path, dst_path)


# Generate documentation.
python_binary = sys.executable
build_binary = r"c:\python25\scripts\sphinx-build-script.py"
build_type = "html"
src_dir = os.path.abspath(os.path.join(directory, "doc-src"))
dst_dir = os.path.abspath(os.path.join(directory, "documentation"))

arguments = [python_binary, build_binary, "-a", "-b", build_type, src_dir, dst_dir]
print "Executing:", arguments
subprocess.call(arguments)
