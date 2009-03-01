#
# This file is part of Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL.
#
#   <http://www.gnu.org/licenses/>.
#

"""
Command-module for sending **keyboard breaks**
==============================================

This module implements the "keyboard break" command.

"""

import pkg_resources
pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")

from dragonfly import (Grammar, MappingRule, Key, Text)


#---------------------------------------------------------------------------
# Defined this module's main rule.

global_rule = MappingRule(
    name="global",
    mapping={
             "keyboard break": Key("c-c"),
            },
    )


#---------------------------------------------------------------------------
# Load the grammar instance and define how to unload it.

global_grammar  = Grammar("global")
global_grammar.add_rule(global_rule)
global_grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global global_grammar
    if global_grammar:  global_grammar.unload()
    global_grammar = None
