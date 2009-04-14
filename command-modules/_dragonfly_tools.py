#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for managing **Dragonfly**
============================================================================

This module manages the configuration files used by other 
active Dragonfly command-modules.  It implements a command 
for easy opening and editing of the configuration files. 
It also monitors the files for modifications, and causes 
the associated command-module to be reloaded if necessary.


Installation
----------------------------------------------------------------------------

If you are using DNS and Natlink, simply place this file in you Natlink 
macros directory.  It will then be automatically loaded by Natlink when 
you next toggle your microphone or restart Natlink.


Commands
----------------------------------------------------------------------------

The following voice commands are available:

Command: **"list configs"**
    Lists the currently available configuration files.
    The output is visible in the *Messages from Python Macros*
    window.

Command: **"edit <config> (config | configuration)"**
    Opens the given configuration file in the default ``*.txt``
    editor.  The *<config>* element should be one of the
    configuration names given by the **"list configs"**
    command.

Command: **"show dragonfly version"**
    Displays the version of the currently active Dragonfly library.

Command: **"update dragonfly version"**
    Updates the Dragonfly library to the newest version available online.

Command: **"reload natlink"**
    Reloads Natlink.

"""

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r81")
except ImportError:
    pass

import os, os.path
from dragonfly import (Grammar, CompoundRule, DictList, DictListRef,
                       MappingRule, Mimic, Key, FocusWindow,
                       Window, Config, Section, Item)


#---------------------------------------------------------------------------
# Set up this module's configuration.

config                   = Config("config manager")
config.lang              = Section("Language section")
config.lang.list_configs = Item("list configs",
                                doc="Command to ...")
config.lang.edit_config  = Item("edit <config> (config | configuration)",
                                doc="Command to ...")
config.lang.show_dragonfly_version = Item("show dragonfly version",
                                doc="Command to ...")
config.lang.update_dragonfly = Item("update dragonfly version",
                                doc="Command to ...")
config.lang.reload_natlink   = Item("reload natlink",
                                doc="Command to ...")
config.load()


#---------------------------------------------------------------------------

config_map = DictList("config_map")


#---------------------------------------------------------------------------

class ConfigManagerGrammar(Grammar):

    def __init__(self):
        Grammar.__init__(self, name="config manager", context=None)

    def _process_begin(self, executable, title, handle):
        configs = Config.get_instances()

        # Check for modified config files, and if found cause reload.
        new_config_map = {}
        for c in configs:
            new_config_map[c.name] = c
            if not os.path.isfile(c.config_path):
                continue
            config_time = os.path.getmtime(c.config_path)
            module_time = os.path.getmtime(c.module_path)
            if config_time >= module_time:
                print "reloading config",c.name
                os.utime(c.module_path, None)

        # Refresh the mapping of config names -> config files.
        config_map.set(new_config_map)

grammar = ConfigManagerGrammar()


#---------------------------------------------------------------------------

class ListConfigsRule(CompoundRule):

    spec = config.lang.list_configs

    def _process_recognition(self, node, extras):
        print "Active configuration files:"
        configs = config_map.keys()
        configs.sort()
        for config in configs:
            print "  - %s" % config

grammar.add_rule(ListConfigsRule())


#---------------------------------------------------------------------------

class EditConfigRule(CompoundRule):

    spec = config.lang.edit_config
    extras = [DictListRef("config", config_map)]

    def _process_recognition(self, node, extras):
        config_instance = extras["config"]
        path = config_instance.config_path
        if not os.path.isfile(path):
            try:
                config_instance.generate_config_file(path)
            except Exception, e:
                self._log.warning("Failed to create new config file %r: %s"
                                  % (path, e))
                return
        os.startfile(path)

grammar.add_rule(EditConfigRule())


#---------------------------------------------------------------------------

class ShowVersionRule(CompoundRule):

    spec = config.lang.show_dragonfly_version

    def _process_recognition(self, node, extras):
        dragonfly_dist = pkg_resources.get_distribution("dragonfly")
        print "Current Dragonfly version:", dragonfly_dist.version
        import dragonfly.engines.engine
        engine = dragonfly.engines.engine.get_engine()
        print "Current language: %r" % engine.language

grammar.add_rule(ShowVersionRule())


#---------------------------------------------------------------------------

class UpdateDragonflyRule(CompoundRule):

    spec = config.lang.update_dragonfly

    def _process_recognition(self, node, extras):
        from pkg_resources import load_entry_point
        import sys
        class Stream(object):
            stream = sys.stdout
            def write(self, data): self.stream.write(data)
            def flush(self): pass
        sys.argv = [""]; sys.stdout = Stream()
#        load_entry_point('setuptools', 'console_scripts', 'easy_install')(["--verbose", "--upgrade", "dragonfly"])
        load_entry_point('setuptools', 'console_scripts', 'easy_install')(["--dry-run", "--upgrade", "dragonfly"])

grammar.add_rule(UpdateDragonflyRule())


#---------------------------------------------------------------------------

class StaticRule(MappingRule):

    mapping = {
               config.lang.reload_natlink: FocusWindow("natspeak", "Messages from Python Macros")
                                            + Key("a-f, r"),
              }

grammar.add_rule(StaticRule())


#---------------------------------------------------------------------------
# Load this module's grammar.

grammar.load()
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
