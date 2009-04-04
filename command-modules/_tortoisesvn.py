# coding=utf-8
#
# (c) Copyright 2008 by Daniel J. Rocco
# Licensed under the Creative Commons Attribution-
#  Noncommercial-Share Alike 3.0 United States License, see
#  <http://creativecommons.org/licenses/by-nc-sa/3.0/us/>
#

"""
Command-module for controlling **TortoiseSVN** from Windows Explorer
============================================================================
    
This module implements various voice-commands for using the
Windows Explorer extensions of the TortoiseSVN subversion client.

(c) Copyright 2008 by Daniel J. Rocco

Licensed under the Creative Commons Attribution-
Noncommercial-Share Alike 3.0 United States License, see
<http://creativecommons.org/licenses/by-nc-sa/3.0/us/>
    
"""

import os.path
import subprocess
import os
import win32gui
import urllib
#from subprocess import Popen

from dragonfly import (Grammar, ConnectionGrammar, AppContext, CompoundRule,
                       Choice, Window, Config, Section, Item)


#---------------------------------------------------------------------------
# Set up this module's configuration.

config                     = Config("TortoiseSVN")
config.tortoisesvn         = Section("TortoiseSVN configuration")
config.tortoisesvn.path    = Item(r'C:\Program Files\TortoiseSVN\bin\TortoiseProc.exe')
config.tortoisesvn.command = Item("(tortoise | subversion) <command> [<predef>]")
config.tortoisesvn.global_command = Item("(tortoise | subversion) <command> <predef>")
config.tortoisesvn.actions = Item({
                                   "add":         "add",
                                   "checkout":    "checkout",
                                   "commit":      "commit",
                                   "revert":      "revert",
                                   "merge":       "merge",
                                   "delete":      "delete",
                                   "diff":        "diff",
                                   "log":         "log",
                                   "import":      "import",
                                   "update":      "update",
                                   "revert":      "revert",
                                   "ignore":      "ignore",
                                   "rename":      "rename",
                                   "properties":  "properties",
                                   "repository":  "repobrowser",
                                   "edit conflict": "conflicteditor",
                                  },
                                 )
config.tortoisesvn.predef  = Item({
                                   "dragonfly | dee fly": r"C:\data\projects\Dragonfly\work dragonfly",
                                  },
                                 )
#config.generate_config_file()
config.load()


#---------------------------------------------------------------------------
# Utility generator function for iterating over COM collections.

def collection_iter(collection):
    for index in xrange(collection.Count):
        yield collection.Item(index)


#---------------------------------------------------------------------------
# This module's grammar for use within Windows Explorer.

class ExplorerGrammar(ConnectionGrammar):

    def __init__(self):
        ConnectionGrammar.__init__(
            self,
            name="Explorer subversion",
            context=AppContext(executable="explorer"),
            app_name="Shell.Application"
           )

    def get_active_explorer(self):
        handle = Window.get_foreground().handle
        for window in collection_iter(self.application.Windows()):
            if window.HWND == handle:
                return window
        self._log.warning("%s: no active explorer." % self)
        return None

    def get_current_directory(self):
        window = self.get_active_explorer()
        path = urllib.unquote(window.LocationURL[8:])
        if path.startswith("file:///"):
            path = path[8:]
        return path

    def get_selected_paths(self):
        window = self.get_active_explorer()
        items = window.Document.SelectedItems()
        paths = []
        for item in collection_iter(items):
            paths.append(item.Path)
        return paths

    def get_selected_filenames(self):
        paths = self.get_selected_paths()
        return [os.path.basename(p) for p in paths]


#---------------------------------------------------------------------------
# Create the rule from which the other rules will be derived.
#  This rule implements the method to execute TortoiseSVN.

class TortoiseRule(CompoundRule):

    def _execute_command(self, path_list, command):
        # Construct arguments and spawn TortoiseSVN.
        path_arg = '/path:"%s"' % str('*'.join(path_list))
        command_arg = "/command:" + command
        os.spawnv(os.P_NOWAIT, config.tortoisesvn.path,
                  [config.tortoisesvn.path, command_arg, path_arg])

        # For some reason the subprocess module gives quote-related errors.
        #Popen([tortoise_path, command_arg, path_arg])


#---------------------------------------------------------------------------
# Create the rule for controlling TortoiseSVN from Windows Explorer.

class ExplorerCommandRule(TortoiseRule):

    spec = config.tortoisesvn.command
    extras = [
              Choice("command", config.tortoisesvn.actions),
             ]
    
    def _process_recognition(self, node, extras):
        selection = self.grammar.get_selected_paths()
        if not selection:
            selection = [self.grammar.get_current_directory()]
        self._execute_command(selection, extras["command"])


#---------------------------------------------------------------------------
# Create the rule for controlling TortoiseSVN from anywhere.

class GlobalCommandRule(TortoiseRule):

    spec = config.tortoisesvn.global_command
    extras = [
              Choice("command", config.tortoisesvn.actions),
              Choice("predef",  config.tortoisesvn.predef),
             ]
    
    def _process_recognition(self, node, extras):
        path_list = [extras["predef"]]
        command = extras["command"]
        self._execute_command(path_list, command)


#---------------------------------------------------------------------------
# Load the grammar instance and define how to unload it.

explorer_grammar = ExplorerGrammar()
explorer_grammar.add_rule(ExplorerCommandRule())
global_grammar = Grammar("TortoiseSVN global")
global_grammar.add_rule(GlobalCommandRule())

explorer_grammar.load()
global_grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global explorer_grammar, global_grammar
    if explorer_grammar:
        explorer_grammar.unload()
        explorer_grammar = None
    if global_grammar:
        global_grammar.unload()
        global_grammar = None
