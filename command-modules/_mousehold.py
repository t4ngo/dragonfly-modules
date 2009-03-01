#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#


import pkg_resources
pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")

import win32con
from ctypes    import (pointer, c_ulong)
from dragonfly import (Grammar, CompoundRule, Choice,
                       Config, Section, Item)
import dragonfly.actions.sendinput as sendinput


#---------------------------------------------------------------------------
# Create the main command rule.

class CommandRule(CompoundRule):

    spec     = "[<button>] mouse <updown>"
    updown   = {
                "down":   0,
                "up":     1,
               }
    buttons  = {
                "left":   0,
                "middle": 1,
                "right":  2,
               }
    extras   = [
                Choice("button", buttons),
                Choice("updown", updown),
               ]

    flag_map = {
                (0,0): win32con.MOUSEEVENTF_LEFTDOWN,
                (0,1): win32con.MOUSEEVENTF_LEFTUP,
                (1,0): win32con.MOUSEEVENTF_MIDDLEDOWN,
                (1,1): win32con.MOUSEEVENTF_MIDDLEUP,
                (2,0): win32con.MOUSEEVENTF_RIGHTDOWN,
                (2,1): win32con.MOUSEEVENTF_RIGHTUP,
               }

    def _process_recognition(self, node, extras):
        updown = extras["updown"]
        if "button" in extras:
            button = extras["button"]
        else:
            button = 0

        key = (button, updown)
        flags = self.flag_map[key]

        input = sendinput.MouseInput(0, 0, 0, flags, 0, pointer(c_ulong(0)))
        array = sendinput.make_input_array([input])
        sendinput.send_input_array(array)


#---------------------------------------------------------------------------
# Create and load this module's grammar.

grammar = Grammar("mouse hold")
grammar.add_rule(CommandRule())
grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
