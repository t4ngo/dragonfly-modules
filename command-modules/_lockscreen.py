#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **screen locking**
=====================================

This module implements a command to lock a computer's screen.
The locking is the same as when the user presses *win-L*.

"""

import pkg_resources
pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")

import ctypes
import natlink
from dragonfly import (Grammar, CompoundRule, Config, Section, Item)


#---------------------------------------------------------------------------
# Set up this module's configuration.

config = Config("lock screen")
config.lang             = Section("Language section")
config.lang.lock_screen = Item("lock screen now now",
                               doc="Command to lock the screen;"
                                   " also puts the microphone to sleep.")
config.load()


#---------------------------------------------------------------------------
# Lock screen rule.

class LockRule(CompoundRule):

    spec = config.lang.lock_screen

    def _process_recognition(self, node, extras):
        self._log.debug("%s: locking screen." % self)

        # Put the microphone to sleep.
        natlink.setMicState("sleeping")

        # Lock screen.
        success = ctypes.windll.user32.LockWorkStation()
        if not success:
            self._log.error("%s: failed to lock screen." % self)


#---------------------------------------------------------------------------
# Create and manage this module's grammar.

grammar = Grammar("lock screen")
grammar.add_rule(LockRule())
grammar.load()
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
