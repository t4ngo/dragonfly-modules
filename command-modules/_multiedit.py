#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for cursor movement and **editing**
============================================================================

This module allows the user to control the cursor and 
efficiently perform multiple text editing actions within a 
single phrase.


Example commands
----------------------------------------------------------------------------

*Note the "/" characters in the examples below are simply 
to help the reader see the different parts of each voice 
command.  They are not present in the actual command and 
should not be spoken.*

Example: **"up 4 / down 1 page / home / space 2"**
   This command will move the cursor up 4 lines, down 1 page,
   move to the beginning of the line, and then insert 2 spaces.

Example: **"left 7 words / backspace 3 / insert hello Cap world"**
   This command will move the cursor left 7 words, then delete
   the 3 characters before the cursor, and finally insert
   the text "hello World".

"""


#---------------------------------------------------------------------------

from dragonfly.all import (Grammar, CompoundRule, MappingRule,
                           Dictation, RuleRef, Repetition,
                           Key, Text, Integer,
                           Config, Section, Item)


#---------------------------------------------------------------------------
# Here we define the keystroke rule.

# This rule maps spoken-forms to actions.  Some of these 
#  include special elements like the number with name "n" 
#  or the dictation with name "text".  This rule is not 
#  exported, but is referenced by other elements later on. 
#  It is derived from MappingRule, so that its "value" when 
#  processing a recognition will be the right side of the 
#  mapping: an action.
class KeystrokeRule(MappingRule):

    exported = False
    mapping  = {
                "up [<n>]":                         Key("up:%(n)d"),
                "down [<n>]":                       Key("down:%(n)d"),
                "left [<n>]":                       Key("left:%(n)d"),
                "right [<n>]":                      Key("right:%(n)d"),
                "page up [<n>]":                    Key("pgup:%(n)d"),
                "page down [<n>]":                  Key("pgdown:%(n)d"),
                "up <n> (page | pages)":            Key("pgup:%(n)d"),
                "down <n> (page | pages)":          Key("pgdown:%(n)d"),
                "left <n> (word | words)":          Key("c-left:%(n)d"),
                "right <n> (word | words)":         Key("c-right:%(n)d"),
                "home":                             Key("home"),
                "end":                              Key("end"),
                "doc home":                         Key("c-home"),
                "doc end":                          Key("c-end"),

                "space [<n>]":                      Key("space:%(n)d"),
                "enter [<n>]":                      Key("enter:%(n)d"),
                "delete [<n>]":                     Key("del:%(n)d"),
                "delete [<n> | this] (line|lines)": Key("home, s-down:%(n)d, del"),
                "backspace [<n>]":                  Key("backspace:%(n)d"),

                "insert <text>":                    Text("%(text)s"),
               }
    extras   = [
                Integer("n", 1, 100),
                Dictation("text"),
               ]
    defaults = {
                "n": 1,
               }

#---------------------------------------------------------------------------
# Here we create an element which is the sequence of keystrokes.

# First we create an element that references the keystroke rule.
keystroke = RuleRef(rule=KeystrokeRule())

# Second we create a repetition of keystroke elements.
#  This element will match anywhere between 1 and 16 repetitions
#  of the keystroke elements.  Note that we give this element
#  the name "sequence" so that it can be used as an extra in
#  the rule definition below.
sequence = Repetition(keystroke, min=1, max=16, name="sequence")


#---------------------------------------------------------------------------
# Here we define the top-level rule which the user can say.

class RepeatRule(CompoundRule):

    # Here we define this rule's spoken-form and special elements.
    spec     = "<sequence> [[[and] repeat [that]] <n> times]"
    extras   = [
                sequence,             # Sequence of actions defined above in.
                Integer("n", 1, 100), # Times to repeat the sequence.
               ]
    defaults = {
                "n": 1,               # Default repeat count.
               }

    # This method gets called when this rule is recognized.
    # Arguments:
    #  - node -- root node of the recognition parse tree.
    #  - extras -- dict of the "extras" special elements:
    #     . extras["sequence"] gives the sequence of actions.
    #     . extras["n"] gives the repeat count.
    def _process_recognition(self, node, extras):
        sequence = extras["sequence"]
        count = extras["n"]
        for i in range(count):
            for action in sequence:
                action.execute()


#---------------------------------------------------------------------------
# Create and load this module's grammar.

grammar = Grammar("multi edit")   # Create this module's grammar.
grammar.add_rule(RepeatRule())    # Add the top-level rule.
grammar.load()                    # Load the grammar.

# Unload function which will be called at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
