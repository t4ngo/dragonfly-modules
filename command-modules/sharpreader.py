#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **SharpReader** feed aggregator
==================================================

This module offers various commands for `SharpReader
<http://www.sharpreader.net/>`_, a free lightweight feed
aggregator.

Commands
--------

Command: **"[feed] address [bar]"**
    Shift focus to the address bar.

Command: **"subscribe [[to] [this] feed]"**
    Add the current feed to the list of subscriptions.

Command: **"paste [feed] address"**
    Paste the contents of the clipboard into the address bar.

Command: **"feeds | feed (list | window | win)"**
    Shift focus to the list of subscribed feeds.

Command: **"newer [<n>]"**
    Move up the list of items.  If *n* is given, move up
    the number of items.

Command: **"older [<n>]"**
    Move down the list of items.  If *n* is given, move down
    the number of items.

"""


#---------------------------------------------------------------------------

from dragonfly.all import (Grammar, AppContext, Rule, MappingRule,
                           Dictation, Alternative, Choice, RuleRef,
                           Key, Text,
                           Config, Section, Item)
from dragonfly.grammar.integer_en  import Integer, Digits


#---------------------------------------------------------------------------
# Create the main command rule.

class CommandRule(MappingRule):

    mapping  = {
                "[feed] address [bar]":                Key("a-d"),
                "subscribe [[to] [this] feed]":        Key("a-u"),
                "paste [feed] address":                Key("a-d, c-v, enter"),
                "feeds | feed (list | window | win)":  Key("a-d, tab:2, s-tab"),
                "down [<n>] (feed | feeds)":           Key("a-d, tab:2, s-tab, down:%(n)d"),
                "up [<n>] (feed | feeds)":             Key("a-d, tab:2, s-tab, up:%(n)d"),
                "open [item]":                         Key("a-d, tab:2, c-s"),
                "newer [<n>]":                         Key("a-d, tab:2, up:%(n)d"),
                "older [<n>]":                         Key("a-d, tab:2, down:%(n)d"),
                "mark all [as] read":                  Key("cs-r"),
                "mark all [as] unread":                Key("cs-u"),
                "search [bar]":                        Key("a-s"),
                "search [for] <text>":                 Key("a-s") + Text("%(text)s\n"),
               }
    extras   = [
                Integer("n", 1, 20),
                Dictation("text"),
               ]
    defaults = {
                "n": 1,
               }


#---------------------------------------------------------------------------
# Create and load this module's grammar.

context = AppContext(executable="SharpReader")
grammar = Grammar("sharp reader", context=context)
grammar.add_rule(CommandRule())
grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
