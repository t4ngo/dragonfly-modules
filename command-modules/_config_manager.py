#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
"""


#---------------------------------------------------------------------------

import os, os.path
from dragonfly.all import (Grammar, CompoundRule, DictList, DictListRef,
                           Config, Section, Item)


#---------------------------------------------------------------------------
# Set up this module's configuration.

config                   = Config("config manager")
config.lang              = Section("Language section")
config.lang.list_configs = Item("list configs",
                                doc="Command to ...")
config.lang.edit_config  = Item("edit <config> (config | configuration)",
                                doc="Command to ...")
config.load()


#---------------------------------------------------------------------------

config_map = DictList("config_map")


#---------------------------------------------------------------------------

class ConfigManagerGrammar(Grammar):

    def __init__(self):
        Grammar.__init__(self, name="config manager", context=None)

    def process_begin(self, executable, title, handle):
        configs = Config.get_instances()

        # Check for modified config files, and if found cause reload.
        new_config_map = {}
        for c in configs:
            if not os.path.isfile(c.config_path): continue
            new_config_map[c.name] = c.config_path
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
        path = extras["config"]
        os.startfile(path)

grammar.add_rule(EditConfigRule())


#---------------------------------------------------------------------------
# Load this module's grammar.

grammar.load()
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
