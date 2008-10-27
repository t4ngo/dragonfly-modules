import sys
import os
import subprocess

python_binary = sys.executable
build_binary = r"c:\python25\scripts\sphinx-build-script.py"
build_type = "html"
directory = os.path.dirname(__file__)
src_dir = os.path.abspath(os.path.join(directory, "doc-src"))
dst_dir = os.path.abspath(os.path.join(directory, "documentation"))

arguments = [python_binary, build_binary, "-a", "-b", build_type, src_dir, dst_dir]
print "Executing:", arguments
subprocess.call(arguments)
