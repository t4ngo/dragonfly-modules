#
# This file is part of Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL.
#
#   <http://www.gnu.org/licenses/>.
#

"""
    This module implements the "keyboard break" command.
"""


from dragonfly.grammar.grammar     import Grammar
from dragonfly.grammar.mappingrule import MappingRule
from dragonfly.actions.actions     import Key, Text


#---------------------------------------------------------------------------
# Create this module's grammar and the context under which it'll be active.

global_grammar  = Grammar("global")


#---------------------------------------------------------------------------

global_rule = MappingRule(
    name="global",
    mapping={
             "keyboard break":          Key("c-c"),
            },
    )

# Add the action rule to the grammar instance.
global_grammar.add_rule(global_rule)


#---------------------------------------------------------------------------
# Load the grammar instance and define how to unload it.

global_grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global global_grammar
    if global_grammar:  global_grammar.unload()
    global_grammar = None
