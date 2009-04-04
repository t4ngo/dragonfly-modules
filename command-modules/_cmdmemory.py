#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for command memory
============================================================================

...

"""

import pkg_resources
pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r97")

import inspect

from dragonfly import *


#---------------------------------------------------------------------------
# Set up this module's configuration.

config                       = Config("command memory")
config.lang                  = Section("Language section")
config.lang.playback_last    = Item("(playback | repeat) [last] [<n>] (commands | command | recognitions) [<count> times]")
config.lang.remember         = Item("remember [last] [<n>] (commands | command | recognitions) as <name>")
config.lang.start_recording  = Item("start recording <name>")
config.lang.stop_recording   = Item("stop recording")
config.lang.playback_memory  = Item("<memory> [<count> times]")
config.lang.recall           = Item("recall everything")
config.load()


#---------------------------------------------------------------------------
# Create global dicts for storing command memories.

# Dictionary for storing memories.
memories     = DictList("memories")
memories_ref = DictListRef("memory", memories)

# Recognition observer for retrieving recently heard recognitions.
playback_history = PlaybackHistory()
playback_history.register()

record_history = PlaybackHistory()
record_name = None


#---------------------------------------------------------------------------
# Define this module's main functionality and rule.

def print_history():
    print "history", playback_history

def playback_last(n, count):
    # Retrieve playback-action from recognition observer.
    playback_history.pop()          # Remove playback recognition itself.
    action = playback_history[-n:]  # Retrieve last *n* recognitions.
    action.speed = 10               # Playback at 10x original speed.

    # Playback action.
    playback_history.unregister()   # Don't record playback.
    try:
        for index in range(count):  # Playback *count* times.
            action.execute()        # Perform playback.
    finally:
        playback_history.register() # Continue recording after playback.

def remember(n, name):
    # Retrieve playback-action from recognition observer and store it.
    playback_history.pop()          # Remove playback recognition itself.
    action = playback_history[-n:]  # Retrieve last *n* recognitions.
    action.speed = 10               # Playback at 10x original speed.
    memories[str(name)] = action    # Store playback action.

def start_recording(name):
    global record_name
    record_name = str(name)
    print "record history", record_history
    del record_history[:]
    print "record history", record_history
    record_history.register()

def stop_recording():
    print "record history", record_history
    record_history.unregister()
    if record_history: 
        record_history.pop()        # Remove playback recognition itself.
    action = record_history[:]
    action.speed = 10               # Playback at 10x original speed.
    memories[record_name] = action  # Store playback action.

def playback_memory(count, memory):
    print "playback history", playback_history
    if playback_history: 
        playback_history.pop()      # Remove playback recognition itself.

    # Playback remembered action.
    playback_history.unregister()   # Don't record playback.
    try:
        for index in range(count):  # Playback *count* times.
            memory.execute()        # Perform playback.
    finally:
        playback_history.register() # Continue recording after playback.


class PlaybackRule(MappingRule):

    mapping  = {  # Spoken form   ->  ->  ->  action
                config.lang.playback_last:    Function(playback_last),
                config.lang.start_recording:  Function(start_recording),
                config.lang.stop_recording:   Function(stop_recording),
                config.lang.remember:         Function(remember),
                config.lang.playback_memory:  Function(playback_memory),
                config.lang.recall:           Function(print_history),
               }
    extras   = [
                IntegerRef("n", 1, 100),      # *n* designates the number
                                              #  of recent recognitions.
                IntegerRef("count", 1, 100),  # *count* designates how
                                              #  many times to repeat
                                              #  playback.
                Dictation("name"),            # *name* is used when
                                              #   Naming new memories.
                memories_ref,                 # This is the list of
                                              #  already-remembered
                                              #  memories; its name is
                                              #  *memory*.
               ]
    defaults = {
                "n": 1,
                "count": 1,
               }


#---------------------------------------------------------------------------
# Create and load this module's grammar.

grammar = Grammar("command memory")     # Create this module's grammar.
grammar.add_rule(PlaybackRule())        # Add the top-level rule.
grammar.load()                          # Load the grammar.

# Unload function which will be called at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
